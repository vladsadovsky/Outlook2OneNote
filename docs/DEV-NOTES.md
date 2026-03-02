# DEV Notes

## 2026-03-01 — OneNote header layout (Windows 11 app)

### Problem
- Exported pages show very narrow metadata header columns in the Windows 11 OneNote app.

### Approaches tried
- Table layout with fixed pixel widths.
- Table layout with percentage widths and explicit width attributes.

### Outcome
- Columns still rendered narrow in the Windows 11 app.

### Next step
- Replace the metadata table with a non-table layout (stacked label/value or div-based two-column layout) to avoid OneNote table width constraints.

## 2026-03-02 — Dual-account support status and estimate

### Current state
- The codebase is not designed as MSA-only.
- Account policy supports `all`, `msa-only`, and `entra-only` via `VITE_ACCOUNT_TYPE_POLICY`.
- Current runtime/testing is MSA-biased due to personal-account fallback and messaging paths.
- Enterprise (Entra) support appears architecturally possible but not yet fully validated end-to-end.

### Estimate to fully enable and sign off both MSA + Entra
- Core engineering hardening: **3–5 dev days**
- Cross-tenant/manual validation (MSA + Entra + policy scenarios): **3–5 QA days**
- Practical total to confident sign-off: **~1–2 weeks**

### Scope assumed in estimate
- Auth behavior consistency across account types (silent/popup + policy handling)
- Enterprise-safe UX/error messaging (remove MSA-only assumptions)
- Full test matrix execution for both account types
- Final readiness checks already tracked in `docs/TASKS.md`

## 2026-03-02 — Delivery estimate for remaining pre-Phase-10 scope

### Task-level estimate
- `T-804` Ribbon button + task pane open command: **0.5–1.5 dev days**
- `T-805` Responsive layout for narrow task pane: **0.5–1 dev day**
- Phase 9 (`T-901`..`T-907`) total: **5–10 days**

### Combined estimate before starting Phase 10
- `T-804` + `T-805` + all of Phase 9: **~6–12 working days** (~1.5–2.5 weeks)

### Estimate assumptions
- No major auth/model change introduced during T-OBO decision
- Existing manual test plan remains the acceptance baseline
- AppSource policy checks do not require backend architecture changes beyond currently tracked scope

## 2026-03-02 — T-OBO architectural decision definition

### What decision is required
Decide whether Outlook2OneNote can remain a pure SPA token model (`Office.auth` login hint + MSAL delegated token in client), or must introduce a backend **On-Behalf-Of (OBO)** exchange for AppSource/enterprise readiness.

### Decision options
1. **No OBO backend (keep SPA-only)**
	- Keep current client-side Graph delegated flow.
	- Lowest implementation effort.
	- Risk: policy/certification constraints may reject this for target enterprise scenarios.

2. **Add OBO backend (recommended if required by policy/certification)**
	- Introduce minimal backend (e.g., Azure Function) for secure token exchange.
	- Stronger enterprise security posture and clearer separation of trust boundaries.
	- Higher implementation + operations overhead.

### Decision criteria
- AppSource certification requirements for Outlook add-ins in target deployment model
- Enterprise customer security expectations (Conditional Access, token handling boundaries)
- Operational complexity tolerance (hosting, monitoring, secret management)

### Required output artifact
- Record decision as an ADR in `TASK-LOG.md` (selected option, rationale, consequences, rollback path).
- If OBO selected: add concrete backend implementation tasks before submission-phase execution.
