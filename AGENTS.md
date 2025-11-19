## Agent Runtime Rules
### Purpose
- Operate solely through natural-language instructions; the user does not run commands or click UI on the agent's behalf.
- Improve workflows iteratively using local tooling and shared documentation.
- Playbooks are essential: short natural-language phrases/aliases that map to intents. The agent must check playbooks **first** on every task(01-system/docs/agents/PLAYBOOKS.md)
- Preserve long-term memory so future sessions stay fast and consistent.

### Invariants (Must Always Hold)
- Natural-language only; ask at most once for missing critical inputs.
- Local-first & reuse-first: prefer existing tools, playbooks, and prompts before inventing new flows.
- Minimal change & fast feedback: take the smallest viable next step and surface results quickly.
- Documentation via Lean Logflow: choose the smallest logging mode that fits the work block.
- Privacy & safety: load secrets from `01-system/configs/apis/API-Keys.md` only when needed and never log them.
- All artifacts live under `03-outputs/<tool>/...`; cite relative paths when reporting.
- `registry.yaml` is the authoritative source for tools; `docs/prompts/INDEX.md` is canonical for curated prompts.
- User-facing docs must remain in **Traditional Chineseï¼ˆç¹é«”ä¸­æ–‡ï¼‰**.

### Startup Checklist (Every Session)
1. Read `STATE.md` for the latest scaffold audit and standing next steps.
2. If filesystem structure changed or the audit is stale, run the full scaffold verification; otherwise perform the quick check (verify sentinel files such as `AGENTS.md`, `registry.yaml`, `03-outputs/README.md`, and key tool folders).
3. Load required secrets from `01-system/configs/apis/API-Keys.md` only when a task needs them.
4. Skim `AGENTS.md` for recent changes; when material updates exist, review `PLAYBOOKS.md`, `TOOLS.md`, `SYSTEM_MEMORY.md`, `docs/prompts/INDEX.md`, and `docs/user/INDEX.md` for additional context.
5. Confirm operating constraints (sandboxing, network, approvals) and decide whether to stay in Execution Mode or request Build Mode.

### Modes
- **Execution Mode (default):** Use existing playbooks, tools, and curated prompts. Do not create new assets without explicit approval.
- **Build Mode (on approval):** Create or modify tools/prompts only after the user green-lights the work. Follow the corresponding flow, return to Execution Mode once complete.

### Repository Layout (Live â€” keep updated)
> Maintain this tree as a living snapshot. Update it whenever structure changes or during startup if drift is detected.
```
/
|- 01-system/
|  |- configs/{env.example, apis/{README.md, API-Keys.md}, tools/registry.yaml}
|  |- docs/
|  |  |- agents/{PLAYBOOKS.md, TOOLS.md, TROUBLESHOOTING.md, SYSTEM_MEMORY.md, STATE.md, BOOTSTRAP.md, memory/YYYY-MM.md}
|  |  |- prompts/{README.md, INDEX.md, examples/prompt-template.md, prompt-*.md, collections/...}
|  |  \- user/{README.md, INDEX.md, tools/...}  # Traditional Chinese
|  \- tools/{ops/, llms/, stt/, _categories-README.md}
|- 02-inputs/{downloads/, snapshots/}
|- 03-outputs/{README.md, <tool>/}
|- tool/
|  |- run_remittance_today.ps1, run_invoices_today.ps1, download_invoices_today.ps1, cleanup_invoices_today.ps1, rename_amount_in_folder.ps1, download_yourremittance.py, migrate_to_store_date.ps1, publish_to_github.ps1, install_portable_git.ps1, install_poppler.ps1
\- tools/
   |- git/  # PortableGit runtime
   \- poppler/
   
   (repo root also includes AGENTS.md, INITIAL_SYSTEM_PROMPT.md, Makefile, supplier_map.json, temp-*.txt)
```




### Outputs (Single Source)
- Store every artifact under `03-outputs/<tool>/`. Use descriptive tool or workflow slugs (`report-writer`, `image-cleanup`, etc.).
- Within each tool folder, organize by run as needed (e.g., timestamps, `intermediate/`, `final/`). Apply one scheme consistently and document exceptions in the run summary.
- Transient downloads belong in `03-outputs/<tool>/downloads/` and must be moved or cited before finishing the task.
- Reference outputs with relative paths in the final message and in `SYSTEM_MEMORY.md` entries.

### Where Things Live
- **Playbooks:** `01-system/docs/agents/PLAYBOOKS.md` â€” first stop for mapping phrases to intents.
- **Prompts Library:** `01-system/docs/prompts/` â€” shared, curated prompts indexed in `INDEX.md` (keep metadata current).
- **Tools:** `01-system/tools/<category>/...` with authoritative registration in `registry.yaml`.
- **Tool index (human-readable):** `01-system/docs/agents/TOOLS.md` mirrors the registry for readers.
- **User documentationï¼ˆç¹é«”ä¸­æ–‡ï¼‰:** `01-system/docs/user/INDEX.md` plus `docs/user/tools/<tool>.md` per asset.
- **Memory & State:** `SYSTEM_MEMORY.md` (canonical log), `memory/YYYY-MM.md` (mirrors), `STATE.md` (phase, next steps, scaffold audit).
- **Troubleshooting:** `01-system/docs/agents/TROUBLESHOOTING.md` collects reproducible fixes and escalation paths.

### Execution Mode â€” Operating Procedure
- Resolve intent via playbooks before planning from scratch; clarify once if ambiguous.
- Prefer registered tools and indexed prompts. When multiple assets fit, choose the safest/local option and cite the prompt ID/version in reports.
- Execute the smallest viable step, writing all artifacts to `03-outputs/<tool>/...`.
- Capture key command outputs (summaries, not raw logs) and call out paths in the final response.
- After each work block, apply Lean Logflow (see below) â€” typically a standard run â€” updating `SYSTEM_MEMORY.md` and `STATE.md` only when the triggers apply.

### Build Mode Flow (Tools & Prompts)
1. **Spec (1â€“3 bullets):** name, category, inputs/outputs, side effects; for prompts add model/provider, variables, guardrails.
2. **Scaffold:**
   - Tool wrappers live under `01-system/tools/<category>/<tool-name>/` and default to `03-outputs/<tool-name>/...`.
   - Prompts use `01-system/docs/prompts/prompt-<domain>-<intent>.md` (template below).
3. **Register/Index:** update `registry.yaml` for tools and `docs/prompts/INDEX.md` for prompts immediately.
4. **Smoke test:** run a minimal check; store artifacts under `03-outputs/<tool-name>/tests/` or similar.
5. **Docs updateï¼ˆå«ç¹é«”ä¸­æ–‡ï¼‰:** refresh `TOOLS.md`, `PLAYBOOKS.md`, `docs/user/tools/<tool>.md`, `docs/user/INDEX.md`, and note new assets in `SYSTEM_MEMORY.md`/`STATE.md`. Update the live repository layout if structure changed.
6. **Return to Execution Mode** once the asset is ready.

### Template â€” `01-system/docs/user/tools/<tool-name>.md`ï¼ˆè«‹ä»¥ç¹é«”ä¸­æ–‡æ’°å¯«ï¼‰
```md
# <å·¥å…·åç¨±>
**é¡žåˆ¥**ï¼š<llms|stt|ops|â€¦>
**ç‰ˆæœ¬**ï¼šv0.1 ï¼ˆæ›´æ–°æ—¥æœŸï¼šYYYY-MM-DDï¼‰

## èƒ½åŠ›ç¸½è¦½
- é€™å€‹å·¥å…·å¯ä»¥åšä»€éº¼ï¼ˆé‡é»žæ¢åˆ—ï¼‰ã€‚

## åƒæ•¸èªªæ˜Ž
- `param1`ï¼šç”¨é€”ã€åž‹åˆ¥ã€é è¨­å€¼èˆ‡ç¯„ä¾‹ã€‚
- `param2`ï¼šâ€¦â€¦

## å¸¸è¦‹ç”¨æ³•ï¼ˆé€æ­¥ï¼‰
1. æ­¥é©Ÿä¸€ï¼šâ€¦â€¦
2. æ­¥é©ŸäºŒï¼šâ€¦â€¦
3. æ­¥é©Ÿä¸‰ï¼šâ€¦â€¦

## ç¯„ä¾‹
- **å¿«é€Ÿç¯„ä¾‹**ï¼šä½¿ç”¨æ­¤çŸ­èªžï¼šã€Œâ€¦â€¦ã€â†’ ç”¢å‡ºæ–¼ `03-outputs/<tool-name>/...`
- **é€²éšŽç¯„ä¾‹**ï¼šâ€¦â€¦

## è¼¸å…¥ / è¼¸å‡ºè·¯å¾‘
- è¼¸å…¥ä¾†æºï¼š`02-inputs/...`
- ç”¢å‡ºä½ç½®ï¼š`03-outputs/<tool-name>/...`

## é¢¨éšªèˆ‡æ¬Šé™
- å¯èƒ½çš„å‰¯ä½œç”¨èˆ‡éœ€è¦çš„æ¬Šé™ï¼›é«˜é¢¨éšªå‹•ä½œéœ€å†ç¢ºèªã€‚

## æ•…éšœæŽ’é™¤
- å¸¸è¦‹éŒ¯èª¤èˆ‡è§£æ³•ï¼ˆé€£çµ `01-system/docs/agents/TROUBLESHOOTING.md` ç›¸é—œæ¢ç›®ï¼‰ã€‚

## ç‰ˆæœ¬èˆ‡æ›´æ–°ç´€éŒ„
- v0.1ï¼ˆYYYY-MM-DDï¼‰ï¼šåˆç‰ˆã€‚
```

### Template â€” `01-system/docs/prompts/prompt-<domain>-<intent>.md`
```md
---
id: prompt-<domain>-<intent>-v1
title: <Concise title>
summary: <purpose and when to use>
model: <openai:gpt-4o|anthropic:claude-3.5|google:gemini-1.5|generic>
owner: <user|agent|team>
version: v1
last_updated: YYYY-MM-DD
tags: [<domain>, <intent>, <safety>]
variables:
  - name: <var_name>
    description: <what it is>
    required: true|false
safety:
  constraints:
    - <e.g., no PII, no destructive ops>
  escalation:
    - <when to ask the user for confirmation>
---

## Usage
- When to use: <guidance>
- Invocation notes: <model quirks, rate limits>
- Expected outputs: <format, quality bar, target path under 03-outputs/<tool>/>

## Prompt
<Write the prompt body here. Use {{variables}} for substitutions.>

## Examples
- Input: <â€¦> â†’ Output: <â€¦>

## Change-log
- v1 (YYYY-MM-DD): Initial version.
```

### Prompts â€” Library & Authoring Rules
- `docs/prompts/INDEX.md` is authoritative for discovery; include ID, model/provider, owner, last update, tags, variables, and safety level.
- Prompts are assets used by LLM-capable tools; keep them atomic and composable.
- When editing a prompt, bump its `version`, update the index metadata, and record the change via Lean Logflow.
- Reference prompts by `id` and `version` in reports.

### Playbooks â€” Authoring Rules (Essential)
- Map phrases/aliases â†’ intent â†’ steps â†’ expected outputs (`03-outputs/<tool>/...`).
- Keep entries explicit and reusable; consolidate overlapping steps instead of duplicating.
- Confirm the required tools/prompts exist (or request Build Mode) before finalizing updates.
- Always attempt playbook matching before free-form planning.

## Lean Logflow (Self-Update Rules)
### Step 1 â€” Classify the work
- **Micro run:** single-step, no durable artifact â†’ answer and stop. Skip DocSync.
- **Standard run:** default for multi-step work or when artifacts exist â†’ you will create one Lean Logflow entry.
- **Milestone run:** rare, repo-wide or hand-off events â†’ same entry format, with richer context if needed.

### Step 2 â€” Minimum DocSync
- Append one line to `SYSTEM_MEMORY.md` using `YYYY-MM-DD â€” Title :: change | impact | artifacts` (include relative paths under `03-outputs/<tool>/...`).
- Mirror the entry to `memory/YYYY-MM.md` only when the month changes or the user explicitly requests a digest.
- If no triggers fire in Step 3, stop here.

### Step 3 â€” Deterministic Triggers (run only when true)
1. **Execution state changed?** â†’ Update `STATE.md` when the phase shifts, standing next steps differ, or the user requests a refresh. Keep it to month/phase, one-sentence summary referencing the matching `SYSTEM_MEMORY.md` line, and the current next steps.
2. **Assets moved or added?** â†’ When you add/modify a tool wrapper or prompt file, update in the same work block: `registry.yaml` â†’ `docs/agents/TOOLS.md` â†’ `docs/prompts/INDEX.md` (for prompts) â†’ related playbook entries â†’ user docs in ç¹é«”ä¸­æ–‡ (as applicable).
3. **Playbook intent changed without new tooling?** â†’ Update `PLAYBOOKS.md` and cite the prompts/tools used.
4. **New troubleshooting knowledge?** â†’ Append to `docs/agents/TROUBLESHOOTING.md` with the fix and escalation guidance.
5. **Repository layout changed?** â†’ Refresh the live tree in `AGENTS.md`.
6. **User asked for anything else?** â†’ Honor explicit instructions (e.g., regenerate a digest or status).

All triggered updates should happen in the same work block as the change so DocSync stays lean and atomic. If none of the conditions apply, no further documentation updates are required.

## Tool Discovery & Registry Rules
- Never invoke unregistered tools. If a wrapper exists without a registry entry, propose registering before use.
- Keep registry changes and code updates atomic; do not leave tools half-registered.
- Prompts stay indexed in `docs/prompts/INDEX.md`; do not treat them as tools.

## Security & Keys
- Load only the secrets you need from `API-Keys.md` and avoid logging values.
- Apply least privilege; request confirmation before high-impact or destructive operations.
- Follow prompt safety constraints and escalate when required.

## Reporting
- Summaries must include what ran, key decisions, and cited artifact paths under `03-outputs/<tool>/...`.
- Mention prompt IDs/versions used for LLM steps.
- For errors, provide probable cause and a minimal recovery step.
- For long operations, share concise progress notes without spamming.

## Ask-Once Checklist
- Missing env vars or secrets.
- Preferences affecting model/provider or tool choice.
- Ambiguity about the intent or target tool folder under `03-outputs/`.

## Change Management
- Keep changes small and reversible; propose larger shifts before acting.
- Do not modify this canonical spec unless the user explicitly instructs you to.
- Every approved change must be reflected through Lean Logflow and the relevant docs listed above.

---

**End of canonical spec.**



