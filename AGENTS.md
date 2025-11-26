## Agent Runtime Rules
### Purpose
- Operate solely through natural-language instructions; the user does not run commands or click UI on the agent's behalf.
- Improve workflows iteratively using local tooling and shared documentation.
- Playbooks are essential: check `01-system/docs/agents/PLAYBOOKS.md` first on every task.
- Preserve long-term memory so future sessions stay fast and consistent.

### Invariants (Must Always Hold)
- Natural-language only; ask at most once for missing critical inputs.
- Local-first & reuse-first: prefer existing tools, playbooks, and prompts before inventing new flows.
- Minimal change & fast feedback: take the smallest viable next step and surface results quickly.
- Documentation via Lean Logflow: choose the smallest logging mode that fits the work block.
- Privacy & safety: load secrets from `01-system/configs/apis/API-Keys.md` only when needed and never log them.
- All artifacts live under `03-outputs/<tool>/...`; cite relative paths when reporting.
- `registry.yaml` is the authoritative source for tools; `docs/prompts/INDEX.md` is canonical for curated prompts.
- User-facing docs should be readable and ungarbled (default to English unless otherwise requested).

### Startup Checklist (Every Session)
1. Read `STATE.md` for the latest scaffold audit and standing next steps.
2. If filesystem structure changed or the audit is stale, run the full scaffold verification; otherwise perform the quick check (verify sentinel files such as `AGENTS.md`, `registry.yaml`, `03-outputs/README.md`, and key tool folders).
3. Load required secrets from `01-system/configs/apis/API-Keys.md` only when a task needs them.
4. Skim `AGENTS.md` for recent changes; when material updates exist, review `PLAYBOOKS.md`, `TOOLS.md`, `SYSTEM_MEMORY.md`, `docs/prompts/INDEX.md`, and `docs/user/INDEX.md` for additional context.
5. Confirm operating constraints (sandboxing, network, approvals) and decide whether to stay in Execution Mode or request Build Mode.

### Modes
- **Execution Mode (default):** Use existing playbooks, tools, and curated prompts. Do not create new assets without explicit approval.
- **Build Mode (on approval):** Create or modify tools/prompts only after the user green-lights the work. Follow the corresponding flow, return to Execution Mode once complete.

### Repository Layout (Live - keep updated)
> Maintain this tree as a living snapshot. Update it whenever structure changes or during startup if drift is detected.
```
/
|- 01-system/
|  |- configs/{env.example, apis/{README.md, API-Keys.md}, tools/registry.yaml}
|  |- docs/
|  |  |- agents/{PLAYBOOKS.md, TOOLS.md, TROUBLESHOOTING.md, SYSTEM_MEMORY.md, STATE.md, BOOTSTRAP.md, memory/YYYY-MM.md}
|  |  |- prompts/{README.md, INDEX.md, examples/prompt-template.md, prompt-*.md, collections/...}
|  |  \- user/{README.md, INDEX.md, tools/...}
|  \- tools/
|     |- ops/
|     |  |- remittance-runner/{run.ps1, run_remittance_today.ps1, download_yourremittance.py, convert_msg_to_pdf.py}
|     |  |- invoices-runner/{run.ps1, run_invoices_today.ps1, download_invoices_today.ps1, cleanup_invoices_today.ps1, supplier_map.json}
|     |  |- remit-rename-amount/{run.ps1, rename_amount_in_folder.ps1}
|     |  |- migrate-store-date/{run.ps1, migrate_to_store_date.ps1}
|     |  \- maintenance/{install_portable_git.ps1, install_poppler.ps1, publish_to_github.ps1}
|     |- runtimes/{git/, poppler/poppler-25.07.0/}
|     |- llms/
|     \- _categories-README.md
|- 02-inputs/{downloads/, snapshots/}
|- 03-outputs/{README.md, remittance-runner/, invoices-runner/}

   (repo root also includes .vscode/, AGENTS.md, INITIAL_SYSTEM_PROMPT.md, temp-*.txt)
```

### Outputs (Single Source)
- Store every artifact under `03-outputs/<tool>/`. Use descriptive tool or workflow slugs (`report-writer`, `image-cleanup`, etc.).
- Within each tool folder, organize by run as needed (e.g., timestamps, `intermediate/`, `final/`). Apply one scheme consistently and document exceptions in the run summary.
- Transient downloads belong in `03-outputs/<tool>/downloads/` and must be moved or cited before finishing the task.
- Reference outputs with relative paths in the final message and in `SYSTEM_MEMORY.md` entries.
- Remittance runner runs may include per-date `intermediate/msg-html`, `intermediate/msg-pdf`, and `intermediate/msg-src` folders for converted email artifacts; store folders should keep only final PDFs.

### Where Things Live
- **Playbooks:** `01-system/docs/agents/PLAYBOOKS.md` - first stop for mapping phrases to intents.
- **Prompts Library:** `01-system/docs/prompts/` - shared, curated prompts indexed in `INDEX.md` (keep metadata current).
- **Tools:** `01-system/tools/<category>/...` with authoritative registration in `registry.yaml`.
- **Tool index (human-readable):** `01-system/docs/agents/TOOLS.md` mirrors the registry for readers.
- **User documentation:** `01-system/docs/user/INDEX.md` plus `docs/user/tools/<tool>.md` per asset.
- **Memory & State:** `SYSTEM_MEMORY.md` (canonical log), `memory/YYYY-MM.md` (mirrors), `STATE.md` (phase, next steps, scaffold audit).
- **Troubleshooting:** `01-system/docs/agents/TROUBLESHOOTING.md` collects reproducible fixes and escalation paths.

### Execution Mode - Operating Procedure
- Resolve intent via playbooks before planning from scratch; clarify once if ambiguous.
- Prefer registered tools and indexed prompts. When multiple assets fit, choose the safest/local option and cite the prompt ID/version in reports.
- Execute the smallest viable step, writing all artifacts to `03-outputs/<tool>/...`.
- Capture key command outputs (summaries, not raw logs) and call out paths in the final response.
- After each work block, apply Lean Logflow (see below) - typically a standard run - updating `SYSTEM_MEMORY.md` and `STATE.md` only when the triggers apply.

### Build Mode Flow (Tools & Prompts)
1. **Spec (1-3 bullets):** name, category, inputs/outputs, side effects; for prompts add model/provider, variables, guardrails.
2. **Scaffold:**
   - Tool wrappers live under `01-system/tools/<category>/<tool-name>/` and default to `03-outputs/<tool-name>/...`.
   - Prompts use `01-system/docs/prompts/prompt-<domain>-<intent>.md` (template below).
3. **Register/Index:** update `registry.yaml` for tools and `docs/prompts/INDEX.md` for prompts immediately.
4. **Smoke test:** run a minimal check; store artifacts under `03-outputs/<tool-name>/tests/` or similar.
5. **Docs update:** refresh `TOOLS.md`, `PLAYBOOKS.md`, `docs/user/tools/<tool>.md`, `docs/user/INDEX.md`, and note new assets in `SYSTEM_MEMORY.md`/`STATE.md`. Update the live repository layout if structure changed.
6. **Return to Execution Mode** once the asset is ready.

### Template - `01-system/docs/user/tools/<tool-name>.md`
```md
# <Tool Name>
**Category**: <llms|stt|ops|...>
**Version**: v0.1 (Updated: YYYY-MM-DD)

## Capabilities
- What this tool can do (bullet points).

## Parameters
- `param1`: purpose, type, default, example.
- `param2`: ...

## Typical Usage (step-by-step)
1. Step one: ...
2. Step two: ...
3. Step three: ...

## Examples
- **Quick example**: Phrase "..." -> outputs to `03-outputs/<tool-name>/...`
- **Advanced example**: ...

## Input / Output Paths
- Input: `02-inputs/...`
- Output: `03-outputs/<tool-name>/...`

## Risks & Permissions
- Possible side effects/required permissions; confirm high-risk actions.

## Troubleshooting
- Common errors and fixes (link `01-system/docs/agents/TROUBLESHOOTING.md`).

## Versions
- v0.1 (YYYY-MM-DD): Initial version.
```

### Template - `01-system/docs/prompts/prompt-<domain>-<intent>.md`
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
- Input: <?> -> Output: <?>

## Change-log
- v1 (YYYY-MM-DD): Initial version.
```
