# Outlook2OneNote v2 — Task Implementation Log

Captures decisions, deviations from spec/design, and implementation notes for completed tasks.
High-level status lives in `docs/TASKS.md`. This file records the *why* and *how*.

---

## T-005 — Scaffold TypeScript Project
**Completed:** 2026-02-28

### Stack decisions
| Concern | Planned (DESIGN.md) | Actual | Reason |
|---------|-------------------|--------|--------|
| Build | Webpack 5 | **Vite 5** | User preference; aligns with llm-aggregator; v2 targets Outlook Web + Outlook Desktop New (WebView2) only — no IE11/Classic, so Webpack polyfill chain unnecessary |
| Testing | Jest + ts-jest | **Vitest 2** | Vite-native, Jest-compatible API, no extra config; aligns with llm-aggregator |
| Env vars | `process.env.VITE_*` | **`import.meta.env.VITE_*`** | Standard Vite pattern; aligns with llm-aggregator's `import.meta.env.DEV` |

DESIGN.md updated to reflect Vite and Vitest.

### Files created
**Config (root):**
- `package.json` — v2 dependencies; scripts: dev, build, preview, test, test:watch, lint
- `tsconfig.json` — strict mode, ES2020, React JSX, `@/*` path alias
- `tsconfig.node.json` — for vite.config.ts (ESNext, Node types)
- `vite.config.ts` — multi-page MPA (taskpane, commands, auth callback); HTTPS via `@vitejs/plugin-basic-ssl`; port 3000; Vitest config inline
- `.eslintrc.json` — `@typescript-eslint/recommended` + `office-addins/recommended`; no-semicolons, single-quotes
- `.env.example` — template with `VITE_CLIENT_ID`, `VITE_TENANT_ID`, `VITE_ACCOUNT_TYPE_POLICY`
- `.gitignore` — updated; `.env.example` intentionally unignored

**Fully implemented (no TODOs):**
- `src/utils/logger.ts` — `debugLog` / `debugError` using `import.meta.env.DEV`
- `src/utils/retry.ts` — `withRetry<T>` exponential backoff (identical to llm-aggregator pattern)
- `src/onenote/types.ts` — full type definitions: `EmailMessage`, `EmailAddress`, `AttachmentMetadata`, `Notebook`, `Section`, `Page`

**Stubs (typed signatures + TODO comments):**
- `src/taskpane/taskpane.ts`, `App.tsx`, `views/ExportView.tsx`, `views/SettingsView.tsx`
- `src/taskpane/components/AppHeader.tsx`, `NotebookPicker.tsx`, `ProgressBar.tsx`, `ErrorBanner.tsx`
- `src/commands/commands.ts`
- `src/auth/authService.ts`, `msalConfig.ts`
- `src/services/graphClient.ts`, `mailService.ts`
- `src/onenote/oneNoteService.ts`, `pageBuilder.ts`, `htmlSanitizer.ts`
- `src/settings/settingsService.ts`
- `src/types/settings.ts`, `export.ts`
- `src/utils/dateFormatter.ts`
- `tests/setup.ts`

**HTML entry points:**
- `src/taskpane/taskpane.html`, `src/commands/commands.html`, `src/auth/callback.html`

### Port
Dev server runs on port **3000** (matching existing Azure app redirect URI `https://localhost:3000/auth/callback.html`).

---

## T-007 — JSON Unified Manifest Skeleton
**Completed:** 2026-02-28

### File
`manifest.json` (project root) — Office Add-in JSON Unified Manifest, `devPreview` schema.

### Key choices
- **New GUID** (`c1f2a3b4-...`): separate identity from v1 XML manifest.
- **Dev URLs** point to `https://localhost:3000/...` — update to production URL before AppSource submission.
- **Permission**: `MailboxItem.Read.User` (maps to `ReadItem` — minimum needed to read selected message).
- **Mailbox requirement**: `1.10` minimum (per SPEC.md Section 8).
- **Single ribbon button**: "Export to OneNote" on `mailRead` context, `TabDefault` group. No additional buttons (per AC-05.4).
- **Two runtimes**: `TaskPaneRuntime` (opens taskpane) + `CommandsRuntime` (function commands).
- **Icons**: copied from v1 `assets/` — `icon-16.png`, `icon-32.png`, `icon-80.png`, `icon-128.png`.

### Pre-AppSource checklist
- [ ] Replace dev URLs with production URLs
- [ ] Fill real privacy policy URL + terms URL
- [ ] Confirm manifest GUID is unique (not reusing v1)
- [ ] Run `npx office-addin-manifest validate manifest.json`

---

## T-006 — Azure App Registration
**Completed:** 2026-02-28

### Decision
Create a **new** Azure app registration for v2 (not reusing v1's `a73f5240-e06c-43a3-8328-1fbd80766263`).

Reasons:
- Clean separation: v1 had a client secret in frontend code (security flaw); v2 has none
- v2 redirect URI differs: `https://localhost:3000/auth/callback.html` (v1 used `/src/auth/callback`)
- Easier to manage scopes and settings independently

### Steps
1. [Azure Portal](https://portal.azure.com) → **App registrations** → **New registration**
   - Name: `Outlook2OneNote-v2`
   - Supported account types: **Accounts in any organizational directory and personal Microsoft accounts** (`common` authority)
   - Redirect URI type: **Single-page application (SPA)**
   - Redirect URI value: `https://localhost:3000/auth/callback.html`

2. **API permissions** → Add delegated permissions:
   - `Mail.Read`
   - `Notes.ReadWrite`
   - `offline_access`
   - `User.Read`
   - Do **not** grant admin consent (these are user-delegated)

3. **Authentication** tab:
   - Confirm SPA redirect URI is listed
   - Do **not** create a client secret
   - Implicit grant checkboxes: leave unchecked (PKCE SPA flow doesn't use implicit grant)

4. Copy **Application (client) ID** → create `.env` from `.env.example`:
   ```
   VITE_CLIENT_ID=<paste-client-id-here>
   VITE_TENANT_ID=common
   VITE_ACCOUNT_TYPE_POLICY=all
   ```

5. Before AppSource submission: add production redirect URI:
   `https://www.countinglight.com/outlook2onenote/auth/callback.html`

---

---

## T-101 — msalConfig.ts
**Completed:** 2026-02-28

Already fully implemented in the T-005 scaffold. No changes needed.

**Redirect URI note:** Config uses `${window.location.origin}/src/auth/callback.html`. In Vite dev this resolves to `https://localhost:3000/src/auth/callback.html` — the Azure SPA redirect URI registration must match this exactly.

---

## T-102 — authService.ts
**Completed:** 2026-02-28

### Flow implemented
1. Check sessionStorage cache (with 5-min expiry buffer)
2. Try `Office.auth.getAccessToken()` to extract a login hint (SSO bootstrap token — NOT used as Graph token, T-OBO still pending)
3. `acquireTokenSilent()` with matched account from MSAL cache
4. `acquireTokenPopup()` fallback with `loginHint` pre-filled
5. `AuthError` thrown on all failures with user-readable message

### Decisions
- SSO token used for login hint only (not as Graph token) — avoids OBO dependency for now
- `parseJwtPayload()` helper handles base64url → base64 conversion before `atob()`
- `_msal` singleton initialized once per page lifecycle (MSAL v3 requires `initialize()` before any method)
- Token cached in sessionStorage under `o2on_token` / `o2on_token_expiry` keys (separate from MSAL's own cache)

---

## T-103 — authService unit tests
**Completed:** 2026-02-28

**File:** `tests/auth/authService.test.ts` — 9 tests, all passing.

### Tests
1. Returns cached token without calling MSAL
2. Treats expired cached token as absent, falls through to popup
3. Silent token via SSO login hint + matching MSAL account
4. Silent token when SSO unavailable but MSAL account exists
5. InteractionRequiredAuthError → popup fallback
6. No accounts → popup
7. Popup rejected → throws AuthError
8. Token cached in sessionStorage after acquisition
9. loginHint passed to popup when SSO provides one

### Vitest gotcha
`vi.resetAllMocks()` clears the `PublicClientApplication` factory (not just its calls). Must call `vi.mocked(PublicClientApplication).mockImplementation(() => mockMsal)` in `beforeEach` after each reset.

---

---

## T-201 — settingsService.ts
**Completed:** 2026-02-28

**File:** `src/settings/settingsService.ts` — factory `createSettingsService()`, 9 unit tests passing.

Settings stored under key `o2on_settings` in `roamingSettings`. Draft state in `SettingsView` is not persisted until `settingsService.save()` is called. Immutability enforced via spread copies in `get()`.

---

## T-202–T-205 — AppHeader, SettingsView, keyboard + click-outside handling
**Completed:** 2026-02-28

### AppHeader
Gear button (⚙) calls `onSettingsClick` prop. Styled with Fluent-aligned `app-header` BEM classes.

### SettingsView
Full overlay dialog with:
- Draft state (local `useState`, not persisted until Save)
- `NotebookPicker` sub-component wired to `oneNoteService.listNotebooks()`
- Section name format text input (tokens: `{subject}`, `{date}`)
- Sort order radio group (oldest/newest first)
- Include attachments checkbox
- Preferred link radio group (web / desktop / both)
- Save (async, calls `settingsService.set()` + `settingsService.save()`) and Cancel buttons
- Keyboard: `document.addEventListener('keydown')` in `useEffect` — Escape → cancel, Ctrl+Enter → save
- Click-outside: `onClick` on backdrop `div` ref; dismisses if target is the backdrop itself

---

## T-301 — graphClient.ts
**Completed:** 2026-02-28

**File:** `src/services/graphClient.ts`

### Strategy
- 401: clear cached token (`clearCachedToken()`), re-acquire, retry once (no loop)
- 429: read `Retry-After` header (default 5s), wait, retry once
- Absolute URL support: if path starts with `http`, used verbatim (for `@odata.nextLink` pagination)
- `withRetry` is available for higher-level callers (orchestrator); graphClient does simple one-shot retries

### clearCachedToken export
Added `clearCachedToken()` to `authService.ts` to allow graphClient to evict the session cache on 401.

---

## T-302/T-305/T-306 — oneNoteService.ts
**Completed:** 2026-02-28

**File:** `src/onenote/oneNoteService.ts`

### Interface change
`createPage` signature changed from `(sectionId, message: EmailMessage, title)` (scaffold) to `(sectionId, pageHtml: string, title)` — pre-built HTML is passed in. This decouples `oneNoteService` from `pageBuilder` and from the `includeAttachments` setting. The orchestrator is responsible for building the HTML.

### Graph API mapping
- `GET /me/onenote/notebooks?$select=id,displayName,links` → maps `links.oneNoteWebUrl.href` to `Notebook.webUrl`
- `POST /me/onenote/notebooks/{id}/sections` with `{ displayName }` → returns `{ id, displayName }`
- `POST /me/onenote/sections/{id}/pages` with `Content-Type: text/html` body → maps `links.oneNoteWebUrl/oneNoteClientUrl` to `Page`

3 unit tests covering all three operations (graphClient mocked).

---

## T-401–T-404 — mailService.ts
**Completed:** 2026-02-28

**File:** `src/services/mailService.ts` — factory `createMailService(client: GraphClient)`

### 50-message cap
Fetches `$top=51` (one over limit). If total accumulated messages > 50 after any page, throws `ThreadTooLargeError` with `count` property. No partial export attempted.

### Attachment handling
Uses `$expand=attachments($select=name,size,contentType,isInline)`. Client-side filter `isInline === false` ensures embedded images are not listed as attachments.

### Paging
`@odata.nextLink` passed as absolute path to `client.get()` on subsequent iterations; params set to `undefined` (nextLink already contains encoded query).

7 unit tests (mapper, paging, ThreadTooLargeError, inline attachment filter, ccRecipients).

---

## T-502–T-508 — onenote library
**Completed:** 2026-02-28

### htmlSanitizer.ts
DOM-based sanitization using `DOMParser`. Removes: `<script>`, `<iframe>`, `<object>`, `<embed>`, `<form>`, `<base>`, `<meta>`, `<link>`. Strips: `on*` event attributes, `style` containing `expression()/javascript:/behavior:`, `javascript:` href/src/action, external image `src` (http/https → removed, alt preserved). 11 unit tests.

### pageBuilder.ts
Builds a full OneNote-compatible HTML document. Metadata table: From, To, CC (if present), Date, Subject. Body: `sanitizeHtml(message.bodyHtml)`. Attachments: table with name/type/size columns, gated on `includeAttachments`. All user-generated text HTML-escaped. 11 unit tests.

### oneNoteService boundary rule (T-508)
`src/onenote/` imports only: `@/services/graphClient` (type-only), `./types`, `@/utils/logger`. No imports from `taskpane/`, `auth/`, `commands/`, or Office.js.

---

## T-601–T-605 — exportOrchestrator.ts + ExportView.tsx
**Completed:** 2026-02-28

### exportOrchestrator.ts
New file `src/taskpane/exportOrchestrator.ts`. Takes `ExportOptions` with injected services and `onProgress` callback. Steps:
1. `mailService.getConversationMessages()` → `ThreadTooLargeError` propagated to caller
2. Sort messages per `settings.sortOrder`
3. Build section name using `sectionNameFormat` tokens `{subject}`, `{date}` (capped at 50 chars for OneNote limit)
4. `oneNoteService.createSection()`
5. For each message: `buildPageHtml(msg, includeAttachments)` → `oneNoteService.createPage()`
6. Return `ExportResult`

### ExportView.tsx
Replaces the temporary auth test UI. States: `idle` (shows Export button or Settings prompt), `fetching`/`exporting` (ProgressBar), `error` (ErrorBanner with retry), `success` (links + "Export another"). ThreadTooLargeError has dedicated advisory message.

---

## T-703/T-704 — dateFormatter.ts unit tests
**Completed:** 2026-02-28

5 unit tests covering `formatDateForSection` (YYYY-MM-DD, padding) and `formatDateTimeForPageTitle` (YYYY-MM-DD HH:MM, padding). Tests use format-checking regexes rather than hardcoded UTC dates to be timezone-neutral.

---

## T-704 — retry.ts unit tests
**Completed:** 2026-02-28

4 unit tests: success on first attempt, retry after failure, throw after maxRetries, exponential backoff delay values. The backoff test spies on `setTimeout` directly and invokes the callback immediately to avoid timer dependency.

---

## T-801–T-803 — App.tsx, ErrorBanner, ProgressBar
**Completed:** 2026-02-28

### App.tsx
Props-based service injection: `App` receives `{ settingsService, mailService, oneNoteService }` from `taskpane.ts` (created inside `Office.onReady`). Services passed down to `ExportView` and `SettingsView` as props. No React context needed — single level of prop threading.

### taskpane.ts
Creates services after `Office.onReady()`:
```typescript
const settingsService = createSettingsService()    // needs roamingSettings
const graphClient = createGraphClient()             // lazy token acquisition
const mailService = createMailService(graphClient)
const oneNoteService = createOneNoteService(graphClient)
```
Imports `./taskpane.css` for styles.

### CSS (taskpane.css)
BEM-structured, Fluent/Office-aligned palette (`#0078d4` accent). Key sections: app-header, export-view, progress-bar (CSS spinner), error-banner, notebook-picker, settings-overlay/dialog, settings-field, shared `.btn` variants.

---

*Last updated: 2026-02-28 (Phase 2–8 implementation complete)*
