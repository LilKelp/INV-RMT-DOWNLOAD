import argparse
import re
import sys
from pathlib import Path
from typing import List, Optional, Tuple

import extract_msg
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright


STYLE = """
body { font-family: 'Segoe UI', Arial, sans-serif; font-size: 11pt; color: #222; line-height: 1.5; margin: 24px; }
table { border-collapse: collapse; width: 100%; margin-top: 12px; }
th, td { border: 1px solid #ccc; padding: 6px 8px; text-align: left; font-size: 10pt; vertical-align: top; }
h1,h2,h3 { margin: 0 0 8px 0; }
p { margin: 0 0 8px 0; }
pre { white-space: pre-wrap; }
.alert { background: #fff8e1; border: 1px solid #f1c232; padding: 10px; margin-bottom: 12px; }
"""


def sanitize_component(value: str) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|]', "_", value.strip())
    return cleaned or "Remittance"


def unique_path(folder: Path, base_name: str) -> Path:
    folder.mkdir(parents=True, exist_ok=True)
    target = folder / base_name
    if not target.exists():
        return target
    stem = target.stem
    suffix = target.suffix
    counter = 1
    while True:
        candidate = folder / f"{stem}_{counter}{suffix}"
        if not candidate.exists():
            return candidate
        counter += 1


def clean_body_html(html_body: str) -> str:
    # Remove literal "\n" artifacts that surface in cell content and headers.
    return re.sub(r"\\n\\s*", " ", html_body)


def load_message(msg_path: Path) -> Tuple[str, str]:
    message = extract_msg.Message(str(msg_path))
    try:
        html_body = message.htmlBody
        text_body = message.body or ""
        if isinstance(html_body, bytes):
            html_body = html_body.decode("utf-8", errors="ignore")
        if isinstance(text_body, bytes):
            text_body = text_body.decode("utf-8", errors="ignore")
        return html_body or "", text_body
    finally:
        try:
            message.close()
        except Exception:
            pass


def html_from_message(html_body: str, text_body: str) -> str:
    if html_body:
        body = clean_body_html(html_body)
    else:
        escaped = (
            text_body.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
        )
        body = f"<pre>{escaped}</pre>"
    return f"<html><head><meta charset='utf-8'><style>{STYLE}</style></head><body>{body}</body></html>"


def extract_ref_amount(text: str) -> Tuple[str, str]:
    # Match "EFT Reference Number: 123" or "Payment Reference Number 123"
    # Enforce 'Number' to avoid matching headers like "EFT Reference" followed by address text.
    ref_match = re.search(r"(?:EFT|Payment)\s+Reference\s+Number\s*[:\s]+\s*([0-9A-Za-z-]+)", text, flags=re.IGNORECASE)
    ref = ref_match.group(1).strip() if ref_match else "EFT"
    amt_match = re.search(r"(?im)^Total:\s*([0-9,]+\.\d{2})", text)
    if amt_match:
        amt = amt_match.group(1).strip()
    else:
        nums = re.findall(r"\d{1,3}(?:,\d{3})*(?:\.\d{2})", text)
        amt = nums and sorted(nums, key=lambda n: float(n.replace(",", "")), reverse=True)[0] or "amount"
    return ref, amt


def resolve_intermediate_folders(msg_path: Path) -> Tuple[Path, Path]:
    """
    Route intermediates alongside the date folder so store folders only hold final PDFs.
    Looks for the nearest ancestor named 'files' (expected structure: .../<date>/files/<store>/file.msg).
    """
    intermediate_root: Optional[Path] = None
    for parent in msg_path.parents:
        if parent.name.lower() == "files":
            intermediate_root = parent.parent / "intermediate"
            break
    if intermediate_root is None:
        intermediate_root = msg_path.parent / "intermediate"
    html_dir = intermediate_root / "msg-html"
    fallback_pdf_dir = intermediate_root / "msg-pdf"
    return html_dir, fallback_pdf_dir


def has_amount_token(value: str) -> bool:
    """Treat outputs without digits as placeholders (e.g., 'amount')."""
    return bool(re.search(r"\d", value))


def convert_single(msg_path: Path, page) -> Optional[Path]:
    try:
        html_body, text_body = load_message(msg_path)
        full_html = html_from_message(html_body, text_body)
        soup = BeautifulSoup(full_html, "html.parser")
        ref, amt = extract_ref_amount(soup.get_text("\n"))
        safe_ref = sanitize_component(ref)
        safe_amt = sanitize_component(amt)
        html_dir, fallback_pdf_dir = resolve_intermediate_folders(msg_path)
        html_out = unique_path(html_dir, f"{safe_ref} - {safe_amt}.html")
        target_dir = msg_path.parent if has_amount_token(safe_amt) else fallback_pdf_dir
        pdf_out = unique_path(target_dir, f"{safe_ref} - {safe_amt}.pdf")

        html_out.write_text(full_html, encoding="utf-8")
        page.set_content(full_html)
        page.pdf(path=str(pdf_out), format="A4")
        print(
            f"Converted {msg_path.name} -> {pdf_out.name}"
            + (f" (html stored at {html_out.parent.name})")
        )
        return pdf_out
    except Exception as exc:
        print(f"Failed to convert {msg_path}: {exc}", file=sys.stderr)
        return None


def convert_all(msg_paths: List[Path]) -> None:
    if not msg_paths:
        print("No .msg files provided; nothing to do.")
        return
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        for msg_path in msg_paths:
            convert_single(msg_path, page)
        browser.close()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Convert remittance .msg files to cleaned HTML/PDF with EFT ref naming.")
    parser.add_argument("--msgs", nargs="+", required=True, help="Paths to .msg files to convert.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    paths = [Path(p) for p in args.msgs if Path(p).exists()]
    convert_all(paths)


if __name__ == "__main__":
    main()
