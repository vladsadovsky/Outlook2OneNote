# Outlook2OneNote v2 — Task Tracker

**Last updated:** 2026-02-28 (Phase 2–8 implementation complete)
**Branch:** v2/main
**Spec:** docs/SPEC.md | **Design:** docs/DESIGN.md | **Test Plan:** docs/TEST-PLAN.md

Status key: `[ ]` todo · `[~]` in progress · `[x]` done · `[!]` blocked

---

## Phase 0 — Foundation

- [x] T-001 Create `v2/main` branch from `master` HEAD
- [x] T-002 Write `docs/SPEC.md`
- [x] T-003 Write `docs/DESIGN.md`
- [x] T-004 Write `docs/TASKS.md`
- [x] T-005 Scaffold TypeScript project (Vite, Vitest, ESLint, tsconfig) — see TASK-LOG.md
- [x] T-006 Register Azure app for v2 (client ID, redirect URIs, delegated scopes) — see TASK-LOG.md
- [x] T-007 Create JSON Unified Manifest skeleton — see TASK-LOG.md
- [x] T-008 Write `docs/TEST-PLAN.md` — comprehensive manual test plan (functional, reliability, performance, stability)

## Phase 1 — Auth (US-01)

- [x] T-101 Implement `msalConfig.ts` — MSAL v3 config + `AccountTypePolicy` + authority mapping
- [x] T-102 Implement `authService.ts` — SSO → silent → popup flow
- [x] T-103 Write unit tests for `authService` (MSAL mocked) — 9 tests, all passing
- [~] T-104 Manual test: SSO path works in Outlook Web — popup fallback verified; SSO silent untested (needs prior session)
- [x] T-105 Manual test: MSAL popup fallback works for Entra ID account — verified ✓
- [x] T-106 Manual test: MSAL popup fallback works for MSA account — verified ✓
- [ ] T-107 Manual test: `entra-only` policy blocks MSA login attempt
- [ ] T-108 Manual test: `msa-only` policy blocks Entra login attempt

## Phase 2 — Settings (US-04)

- [x] T-201 Implement `settingsService.ts` with typed `SettingsSchema` (incl. `preferredLink`)
- [x] T-202 Implement `AppHeader.tsx` with ⚙ gear icon toggle
- [x] T-203 Implement `SettingsView.tsx` — overlay dialog (notebook picker, all settings fields)
- [x] T-204 Keyboard handling: Escape to cancel, Ctrl+Enter to save
- [x] T-205 Click-outside-to-dismiss on settings overlay
- [x] T-206 Wire roamingSettings save/load — covered by T-201 implementation
- [x] T-207 Unit tests for `settingsService` — 9 tests, all passing

## Phase 3 — Notebook Selection (US-02)

- [x] T-301 Implement `graphClient.ts` — token injection, 401 retry, 429 backoff
- [x] T-302 Implement `oneNoteService.listNotebooks()` (in `src/onenote/`)
- [x] T-303 Implement `NotebookPicker.tsx` component
- [x] T-304 Persist selected notebook to roamingSettings (via settingsService in SettingsView)
- [~] T-305 Handle "notebook no longer exists" case (AC-02.3) — SettingsView re-loads list on open; full re-selection prompt is manual-test item
- [x] T-306 Unit tests for `oneNoteService` (graphClient mocked) — 3 tests

## Phase 4 — Mail Fetch (US-03 partial)

- [x] T-401 Implement `mailService.getConversationMessages()` with paged response handling
- [x] T-402 Implement 50-message cap: fetch up to 51, throw `ThreadTooLargeError` if exceeded
- [x] T-403 Unit tests for `mailService` including `ThreadTooLargeError` path
- [x] T-404 Unit tests for `mailService` paging (`@odata.nextLink` mock)

## Phase 5 — OneNote Library (src/onenote/)

- [x] T-501 Define `onenote/types.ts` — `EmailMessage`, `EmailAddress`, `AttachmentMetadata`, `Notebook`, `Section`, `Page`
- [x] T-502 Implement `htmlSanitizer.ts` — strip scripts/iframes/event attrs/external images
- [x] T-503 Implement `pageBuilder.ts` — OneNote page HTML from `EmailMessage`
- [x] T-504 Attachment metadata section in page (gated on `includeAttachments`)
- [x] T-505 Implement `oneNoteService.createSection()` and `createPage()`
- [x] T-506 Unit tests for `htmlSanitizer` (malicious/complex HTML inputs) — 11 tests
- [x] T-507 Unit tests for `pageBuilder` — 11 tests
- [x] T-508 Boundary rule confirmed: `src/onenote/` has no runtime imports from outside

## Phase 6 — Export Flow (US-03 complete)

- [x] T-601 Implement export orchestration (`exportOrchestrator.ts`)
- [x] T-602 Progress reporting (message N of M) via `onProgress` callback in orchestrator
- [x] T-603 `ThreadTooLargeError` handling — advisory message with Outlook compression guidance
- [x] T-604 Success state with OneNote link(s) filtered by `preferredLink` setting
- [x] T-605 Error handling + retry — ErrorBanner with retry callback resets to idle state
- [ ] T-606 Integration test: full export flow (fetch mocked at network boundary)

## Phase 7 — Shared Utilities

- [x] T-701 Implement `utils/logger.ts` (`debugLog` / `debugError` — aligned with llm-aggregator)
- [x] T-702 Implement `utils/retry.ts` (`withRetry<T>` — aligned with llm-aggregator)
- [x] T-703 Implement `utils/dateFormatter.ts`
- [x] T-704 Unit tests for `retry.ts` (backoff timing, max retries) — 4 tests

## Phase 8 — UI & Commands

- [x] T-801 `App.tsx` shell with Export / Settings view toggle (services injected as props)
- [x] T-802 `ErrorBanner.tsx` component
- [x] T-803 `ProgressBar.tsx` component
- [ ] T-804 Ribbon button + task pane open command (US-05) — single button only
- [ ] T-805 Responsive layout for narrow task pane

## Phase 9 — Hardening & AppSource Readiness

- [ ] T-901 Run Office Add-in Validator
- [ ] T-902 ESLint clean pass (no `any`, strict mode, no semicolons, single quotes)
- [ ] T-903 Review Graph API scope minimisation (NFR-05)
- [ ] T-904 Performance: task pane load ≤ 3s (NFR-01)
- [ ] T-905 Performance: 20-message export ≤ 30s (NFR-02)
- [ ] T-906 Write `AUTHENTICATION_SETUP.md` for v2 (Azure app registration guide)
- [ ] T-907 Execute full manual test plan from `docs/TEST-PLAN.md`

## Phase 10 — AppSource Submission
*(start only after Phase 9 complete and T-907 manual test pass signed off)*

- [ ] T-1001 Resolve T-OBO: confirm whether OBO backend is required for AppSource certification; add backend tasks if so
- [ ] T-1002 Create Microsoft Partner Center account (if not already registered)
- [ ] T-1003 Write store listing: short description (100 chars), long description (500 chars), keywords
- [ ] T-1004 Capture store screenshots (min 3): task pane open, export in progress, success state
- [ ] T-1005 Publish privacy policy page (required by AppSource)
- [ ] T-1006 Publish support/help page URL
- [ ] T-1007 Review AppSource certification policies for Office Add-ins (checklist pass)
- [ ] T-1008 Submit add-in to Partner Center for AppSource review
- [ ] T-1009 Address any certification feedback from Microsoft review team
- [ ] T-1010 Confirm public listing live on AppSource

---

## Blocked / Decisions Pending

- **T-OBO** — Determine if OBO server-side exchange is required for AppSource, or if direct token use is acceptable. If OBO needed, add Azure Function backend tasks.

---

## Future (not in v2)

- T-F01 Playwright automated E2E test suite
- T-F02 llm-aggregator integration: Outlook2OneNote → aggregator → OneNote pipeline
- T-F03 Extract `utils/retry.ts` + `utils/logger.ts` into shared npm package used by both projects
- T-F04 Outlook Classic / COM add-in support

---

## Completed

- [x] T-001 Create `v2/main` branch
- [x] T-002 `docs/SPEC.md`
- [x] T-003 `docs/DESIGN.md`
- [x] T-004 `docs/TASKS.md`
- [x] T-005 Scaffold TypeScript project (Vite 5, Vitest 2, ESLint 9, tsconfig strict)
- [x] T-007 `manifest.json` — JSON Unified Manifest skeleton
- [x] T-008 `docs/TEST-PLAN.md`
- [x] T-101 `msalConfig.ts`
- [x] T-102 `authService.ts`
- [x] T-103 Unit tests for `authService` (9 tests)
