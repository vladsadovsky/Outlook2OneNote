# Development Notes

Lightweight log for ideas, experiments, and approaches tried during development.
Keep this file short, scannable, and factual.

## Best practices (recommended)

- Date every entry and keep newest first.
- Capture the problem, approaches tried, outcome, and next step.
- Avoid long narratives; link to code or tasks if relevant.
- Record why a path was rejected to prevent rework.
- Move resolved items into TASK-LOG.md if they become permanent decisions.

---

## 2026-03-01 — OneNote header layout (Windows 11 app)

Problem:
- Exported pages show very narrow metadata header columns in the Windows 11 OneNote app.

Approaches tried:
- Table layout with fixed pixel widths.
- Table layout with percentage widths and explicit width attributes.

Outcome:
- Columns still render narrow in the Windows 11 app.

Next step:
- Replace the metadata table with a non-table layout (stacked label/value or div-based two-column layout) to avoid OneNote table width constraints.
