# Copilot Instructions — Outlook2OneNote

## Core workflow
- Follow spec-first sequence: `docs/SPEC.md` → `docs/DESIGN.md` → `docs/TASKS.md` → implementation.
- Keep changes minimal, focused, and aligned with existing architecture.

## Documentation safety (mandatory)
- Never delete or overwrite existing note/log content unless the user explicitly requests replacement.
- For `DEV_NOTES.md`, `docs/DEV-NOTES.md`, `docs/TASK-LOG.md`, and `docs/TASKS.md`, use append-or-merge behavior by default.
- Before editing docs, read current contents and preserve dated entries.
- If a requested edit conflicts with existing content, ask before destructive changes.
- After doc edits, summarize what was added, changed, and removed.
- Do not consolidate/delete similarly named note files without explicit user instruction.

## Validation expectation
- When code changes are made, run the most targeted tests first.
- Do not fix unrelated failures.
