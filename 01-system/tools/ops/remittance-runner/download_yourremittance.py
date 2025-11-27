import argparse
import datetime as dt
import os
import re
import shutil
import subprocess
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple
from urllib.parse import parse_qs, unquote, urlparse

import extract_msg
from playwright.sync_api import sync_playwright
import win32com.client

RUNNER_BASE = Path("03-outputs/remittance-runner")
LOG_SUBDIR = "secure-fetcher"
MAILBOX_MAP = {
    "au-orders@novabio.com": "Australia Orders",
}
DEFAULT_MAILBOX = "Australia Orders"
PASSCODE_TIMEOUT = 180
POLL_INTERVAL = 5
REQUEST_BUTTON = "#btn_send_otp"
VERIFY_BUTTON = "#qwer"
INPUT_SELECTOR = "input[aria-label='Verification passcode']"
PASSCODE_RE = re.compile(r"passcode\s+is\s+(\d{6})", re.IGNORECASE)
DIRECT_URL_RE = re.compile(r"https://[^\s>\"']*yourremittance\.com\.au[^\s>\"']*", re.IGNORECASE)
ENCODED_URL_RE = re.compile(r"a=(https%3a%2f%2fyourremittance\.com\.au[^&]+)", re.IGNORECASE)
AMOUNT_PATTERNS = [
    re.compile(r"TOTAL\s+AMOUNT\s+\$?\s*(-?\d[\d,]*\.\d{2})", re.IGNORECASE),
    re.compile(r"(?:grand\s+total|total\s+amount|amount\s+paid|total\s+paid|net\s+total)\s*\$?\s*(-?\d[\d,]*\.\d{2})", re.IGNORECASE),
    re.compile(r"(?:AUD|NZD)\s*\$?\s*(-?\d[\d,]*\.\d{2})", re.IGNORECASE),
    re.compile(r"Total\s+Paid\s+[\w\s]*\$\s*(-?\d[\d,]*\.\d{2})", re.IGNORECASE),
]
DOC_REF_PATTERNS = [
    re.compile(r"Document\s+Ref[\s\S]{0,120}?No[:\s]+([A-Za-z0-9-]+)", re.IGNORECASE),
    re.compile(r"Reference\s+Number[:\s]+([A-Za-z0-9-]+)", re.IGNORECASE),
    re.compile(r"Our\s+Ref[:\s]+([A-Za-z0-9-]+)", re.IGNORECASE),
]

LOG_DIR: Optional[Path] = None
LOG_FILE: Optional[Path] = None
PDFTOTEXT_PATH: Optional[Path] = None
WORKSPACE_ROOT: Optional[Path] = None


@dataclass
class Job:
    msg_path: Path
    store: str
    date_key: str
    transmission_id: str
    portal_url: str
    recipient: str
    mailbox_name: str

    @property
    def passcode_subject(self) -> str:
        return f"One-time verification passcode for {self.transmission_id}"


def log(msg: str) -> None:
    timestamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {msg}"
    print(line)
    if LOG_DIR is None:
        return
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log_path = LOG_FILE or (LOG_DIR / "session-log.txt")
    with open(log_path, "a", encoding="utf-8") as handle:
        handle.write(line + "\n")


def unique_path(folder: Path, base_name: str) -> Path:
    folder.mkdir(parents=True, exist_ok=True)
    target = folder / base_name
    if not target.exists():
        return target
    stem, ext = os.path.splitext(base_name)
    counter = 1
    while True:
        candidate = folder / f"{stem}_{counter}{ext}"
        if not candidate.exists():
            return candidate
        counter += 1


def is_within_runner(path: Path) -> bool:
    try:
        path.resolve().relative_to(RUNNER_BASE.resolve())
        return True
    except ValueError:
        return False


def sanitize_component(value: str) -> str:
    cleaned = re.sub(r"[\\/:*?\"<>|]", "_", value.strip())
    return cleaned or "Remittance"


def get_workspace_root() -> Path:
    global WORKSPACE_ROOT
    if WORKSPACE_ROOT is not None:
        return WORKSPACE_ROOT
    current = Path(__file__).resolve().parent
    while True:
        if (current / "AGENTS.md").exists():
            WORKSPACE_ROOT = current
            return current
        parent = current.parent
        if parent == current:
            break
        current = parent
    raise RuntimeError("Unable to locate workspace root from download_yourremittance.py")


def get_namespace():
    return win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")


def get_inbox(namespace, mailbox_name: str):
    store = namespace.Folders[mailbox_name]
    return store.Folders["Inbox"]


def snapshot_passcode_ids(namespace, mailbox_name: str, subject_fragment: str, limit: int = 20) -> set:
    known: set = set()
    inbox = get_inbox(namespace, mailbox_name)
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)
    for item in items:
        subject = (item.Subject or "").strip()
        if subject_fragment not in subject:
            continue
        known.add(item.EntryID)
        if len(known) >= limit:
            break
    return known


def wait_for_passcode(job: Job, namespace, known_map: Dict[str, set]) -> str:
    mailbox = job.mailbox_name
    if mailbox not in known_map:
        known_map[mailbox] = snapshot_passcode_ids(namespace, mailbox, job.passcode_subject)
    known_ids = known_map[mailbox]
    inbox = get_inbox(namespace, mailbox)
    deadline = time.time() + PASSCODE_TIMEOUT
    while time.time() < deadline:
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)
        for item in items:
            subject = (item.Subject or "").strip()
            if job.passcode_subject not in subject:
                continue
            entry_id = item.EntryID
            if entry_id in known_ids:
                continue
            body = item.Body or ""
            match = PASSCODE_RE.search(body)
            if match:
                received = item.ReceivedTime
                stamp = received.isoformat()
                log(f"Passcode email received ({stamp}) for transmission {job.transmission_id}")
                known_ids.add(entry_id)
                return match.group(1)
        time.sleep(POLL_INTERVAL)
    raise RuntimeError(f"Timed out waiting for one-time passcode for {job.transmission_id}")


def detect_pdftotext() -> Optional[Path]:
    candidates: List[Path] = []
    which = shutil.which("pdftotext")
    if which:
        candidates.append(Path(which))
    workspace = get_workspace_root()
    repo_candidates = [
        workspace / "01-system" / "tools" / "runtimes" / "poppler" / "poppler-25.07.0" / "Library" / "bin" / "pdftotext.exe",
        workspace / "tools" / "poppler" / "Library" / "bin" / "pdftotext.exe",
        workspace / "tools" / "poppler" / "bin" / "pdftotext.exe",
    ]
    for cand in repo_candidates:
        if cand.exists():
            candidates.append(cand)
    poppler_root = workspace / "tools" / "poppler"
    if poppler_root.exists():
        try:
            hit = next(poppler_root.rglob("pdftotext.exe"))
            candidates.append(hit)
        except StopIteration:
            pass
    for cand in candidates:
        if cand.exists():
            return cand
    return None


def extract_pdf_text(pdf_path: Path) -> str:
    global PDFTOTEXT_PATH
    if PDFTOTEXT_PATH is None:
        PDFTOTEXT_PATH = detect_pdftotext()
    if not PDFTOTEXT_PATH:
        log("pdftotext not available; cannot extract text for metadata parsing.")
        return ""
    try:
        result = subprocess.run(
            [str(PDFTOTEXT_PATH), "-layout", "-nopgbrk", "-q", "-f", "1", "-l", "6", "-enc", "UTF-8", str(pdf_path), "-"],
            check=True,
            capture_output=True,
            text=True,
        )
        return result.stdout
    except subprocess.CalledProcessError as exc:
        log(f"pdftotext failed for {pdf_path.name}: {exc}")
        return ""


def parse_pdf_metadata(pdf_path: Path) -> Tuple[Optional[str], Optional[str]]:
    text = extract_pdf_text(pdf_path)
    if not text:
        return None, None
    amount: Optional[str] = None
    for pattern in AMOUNT_PATTERNS:
        match = pattern.search(text)
        if match:
            amount = match.group(1).replace(",", "")
            break
    doc_ref: Optional[str] = None
    for pattern in DOC_REF_PATTERNS:
        match = pattern.search(text)
        if match:
            doc_ref = match.group(1).strip()
            break
    return doc_ref, amount


def build_target_filename(job: Job, doc_ref: Optional[str], amount: Optional[str], suggested: Optional[str]) -> str:
    base = doc_ref or job.transmission_id
    parts = [sanitize_component(base)]
    if amount:
        parts.append(sanitize_component(amount))
    else:
        stem = Path(suggested).stem if suggested else "Remittance"
        parts.append(sanitize_component(stem))
    return " - ".join(parts) + ".pdf"


def load_processed(log_dir: Path) -> Tuple[set, Path]:
    manifest = log_dir / "processed_ids.txt"
    processed: set = set()
    if manifest.exists():
        for line in manifest.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line:
                processed.add(line)
    return processed, manifest


def record_processed(manifest: Path, transmission_id: str) -> None:
    with open(manifest, "a", encoding="utf-8") as handle:
        handle.write(transmission_id + "\n")


def extract_transmission_id(text: str) -> Optional[str]:
    match = re.search(r"Transmission ID[:\s]+([A-Za-z0-9-]+)", text, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return None


def clean_portal_url(raw: str) -> str:
    cleaned = raw.strip().replace("&amp;", "&")
    # Remove trailing punctuation
    return cleaned.rstrip(").,")


def extract_portal_url(text: str) -> Optional[str]:
    direct = DIRECT_URL_RE.search(text.replace("&amp;", "&"))
    if direct:
        return clean_portal_url(direct.group(0))
    encoded = ENCODED_URL_RE.search(text)
    if encoded:
        return clean_portal_url(unquote(encoded.group(1)))
    parsed = re.search(r"https://nam\d+\.safelinks\.protection\.outlook\.com/[^\s>\"']+", text, re.IGNORECASE)
    if not parsed:
        return None
    candidate = parsed.group(0)
    try:
        query = parse_qs(urlparse(candidate).query)
        raw = query.get("url") or query.get("a")
        if raw:
            candidate = unquote(raw[0])
        inner_query = parse_qs(urlparse(candidate).query)
        inner_raw = inner_query.get("a") or inner_query.get("url")
        if inner_raw:
            candidate = unquote(inner_raw[0])
        if "yourremittance.com.au" in candidate.lower():
            return clean_portal_url(candidate)
    except Exception:
        return None
    return None


def should_process_sender(sender: str) -> bool:
    sender_lower = sender.lower()
    return "yourremittance.com.au" in sender_lower


def discover_jobs(base_dir: Path, date_key: str, stores_filter: Optional[Iterable[str]]) -> List[Job]:
    jobs_by_id: Dict[str, Job] = {}
    stores = []
    if stores_filter:
        stores = [base_dir / store for store in stores_filter]
    else:
        if not base_dir.exists():
            return []
        stores = [p for p in base_dir.iterdir() if p.is_dir()]
    for store_dir in stores:
        if not store_dir.is_dir():
            continue
        targets = []
        dated = store_dir / date_key
        if dated.exists():
            targets.append(dated)
        if store_dir.exists():
            targets.append(store_dir)
        seen = set()
        for target in targets:
            key = str(target.resolve())
            if key in seen:
                continue
            seen.add(key)
            for msg_path in sorted(target.glob("*.msg")):
                job = create_job_from_msg(msg_path, store_dir.name, date_key)
                if job and job.transmission_id not in jobs_by_id:
                    jobs_by_id[job.transmission_id] = job
    return list(jobs_by_id.values())


def create_job_from_msg(msg_path: Path, store: str, date_key: str) -> Optional[Job]:
    message = extract_msg.Message(str(msg_path))
    try:
        sender = (message.sender or "").strip()
        if not should_process_sender(sender):
            return None
        body_parts = []
        if message.body:
            body_parts.append(message.body if isinstance(message.body, str) else message.body.decode("utf-8", "ignore"))
        html = getattr(message, "htmlBody", None)
        if html:
            if isinstance(html, bytes):
                html = html.decode("utf-8", "ignore")
            body_parts.append(html)
        combined = "\n".join(body_parts)
        transmission_id = extract_transmission_id(combined)
        portal_url = extract_portal_url(combined)
        recipient_field = (message.to or "").split(";")[0].strip().lower()
        mailbox = MAILBOX_MAP.get(recipient_field, DEFAULT_MAILBOX)
        if not transmission_id:
            log(f"Skipping {msg_path.name}: Transmission ID not found.")
            return None
        if not portal_url:
            log(f"Skipping {msg_path.name}: secure download link not detected.")
            return None
        return Job(
            msg_path=msg_path,
            store=store,
            date_key=date_key,
            transmission_id=transmission_id,
            portal_url=portal_url,
            recipient=recipient_field,
            mailbox_name=mailbox,
        )
    finally:
        message.close()


def download_for_job(
    job: Job,
    context,
    namespace,
    known_map: Dict[str, set],
    processed_ids: set,
    manifest: Path,
) -> bool:
    log(f"Processing {job.transmission_id} ({job.msg_path.name}) via {job.portal_url}")
    page = context.new_page()
    temp_path: Optional[Path] = None
    run_root = RUNNER_BASE / job.date_key
    store_dir = run_root / "files" / job.store
    downloads_dir = LOG_DIR / "downloads"
    try:
        page.goto(job.portal_url, wait_until="networkidle")
        for _ in range(4):
            try:
                page.wait_for_selector(REQUEST_BUTTON, timeout=20000)
                break
            except Exception:
                log("Passcode request UI not ready, reloading portal page.")
                page.reload(wait_until="networkidle")
        else:
            raise RuntimeError("Unable to locate passcode request button.")
        page.click(REQUEST_BUTTON)
        page.wait_for_selector(INPUT_SELECTOR, timeout=15000)
        passcode = wait_for_passcode(job, namespace, known_map)
        log(f"Applying passcode {passcode} for {job.transmission_id}")
        page.fill(INPUT_SELECTOR, passcode)
        with page.expect_download(timeout=60000) as download_info:
            page.click(VERIFY_BUTTON)
        download = download_info.value
        suggested = download.suggested_filename or f"{job.transmission_id}.pdf"
        temp_path = unique_path(downloads_dir, suggested)
        download.save_as(str(temp_path))
        doc_ref, amount = parse_pdf_metadata(temp_path)
        preferred_name = build_target_filename(job, doc_ref, amount, suggested)
        dest = unique_path(store_dir, preferred_name)
        shutil.move(str(temp_path), str(dest))
        log(f"Saved {dest} (Doc Ref: {doc_ref or 'n/a'}, Amount: {amount or 'n/a'})")
        if is_within_runner(job.msg_path):
            try:
                job.msg_path.unlink(missing_ok=True)
                log(f"Removed placeholder {job.msg_path}")
            except Exception as unlink_error:
                log(f"Failed to remove placeholder {job.msg_path}: {unlink_error}")
        processed_ids.add(job.transmission_id)
        record_processed(manifest, job.transmission_id)
        return True
    except Exception as exc:
        log(f"Error downloading {job.transmission_id}: {exc}")
        if temp_path and temp_path.exists():
            temp_path.unlink()
        return False
    finally:
        page.close()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Download secure remittance PDFs referenced by .msg placeholders.")
    parser.add_argument("--date", help="Target date folder (YYYY-MM-DD). Default: today.")
    parser.add_argument("--stores", nargs="*", help="Explicit store folders to scan (default: all store folders under the base directory).")
    parser.add_argument(
        "--base-dir",
        help="Override the folder that contains store subdirectories with .msg placeholders (defaults to remittance-runner/<date>/files).",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    date_key = args.date or dt.date.today().strftime("%Y-%m-%d")
    global LOG_DIR, LOG_FILE
    run_root = RUNNER_BASE / date_key
    LOG_DIR = run_root / LOG_SUBDIR
    LOG_FILE = LOG_DIR / "session-log.txt"
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    default_base = RUNNER_BASE / date_key / "files"
    base_dir = Path(args.base_dir) if args.base_dir else default_base
    if not base_dir.exists():
        log(f"Base directory not found: {base_dir}")
        return
    processed_ids, manifest = load_processed(LOG_DIR)
    jobs = discover_jobs(base_dir, date_key, args.stores)
    if not jobs:
        log(f"No pending secure remittance placeholders found for {date_key}.")
        return
    namespace = get_namespace()
    known_map: Dict[str, set] = {}
    completed = 0
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True)
        for job in jobs:
            if job.transmission_id in processed_ids:
                log(f"Transmission {job.transmission_id} already processed; skipping {job.msg_path.name}.")
                continue
            try:
                if download_for_job(job, context, namespace, known_map, processed_ids, manifest):
                    completed += 1
            except Exception as exc:
                log(f"Failed to download {job.transmission_id}: {exc}")
        context.close()
        browser.close()
    log(f"Completed {completed} of {len(jobs)} job(s).")


if __name__ == "__main__":
    main()
