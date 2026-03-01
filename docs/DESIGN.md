# Outlook2OneNote v2 — Technical Design

**Version:** 2.1
**Date:** 2026-02-28
**Status:** Draft

---

## 1. Module Architecture

```
src/
  taskpane/
    taskpane.ts          ← Entry point; initialises Office.js, mounts React root
    taskpane.html        ← Shell HTML
    App.tsx              ← Root React component; toggles between Export and Settings views
    views/
      ExportView.tsx     ← Main export UI (notebook picker, export button, progress)
      SettingsView.tsx   ← Settings overlay/dialog (opened via gear icon, not ribbon)
    components/
      AppHeader.tsx      ← Header with title + gear icon button (⚙) for settings
      NotebookPicker.tsx ← Dropdown for notebook selection
      ProgressBar.tsx    ← Export progress indicator
      ErrorBanner.tsx    ← Error display with retry action
  commands/
    commands.ts          ← Ribbon button handler (opens task pane)
    commands.html        ← Commands shell
  auth/
    authService.ts       ← SSO + MSAL orchestration; sole owner of token lifecycle
    msalConfig.ts        ← MSAL PublicClientApplication config + account type policy
  services/
    graphClient.ts       ← Thin wrapper: attaches Bearer token, handles 401 retry + 429 backoff
    mailService.ts       ← Graph calls for reading mail/conversations
  settings/
    settingsService.ts   ← Read/write roamingSettings; typed SettingsSchema

  # ─── Reusable OneNote Library (no Outlook dependencies) ────────────────────
  onenote/
    oneNoteService.ts    ← Graph calls for creating notebooks/sections/pages
    pageBuilder.ts       ← Builds OneNote page HTML from a generic EmailMessage
    htmlSanitizer.ts     ← Best-effort HTML sanitisation for OneNote OOXML
    types.ts             ← Notebook, Section, Page, EmailMessage interfaces
  # ───────────────────────────────────────────────────────────────────────────

  utils/
    dateFormatter.ts     ← Consistent date/time formatting
    logger.ts            ← Conditional dev-only logging (debugLog / debugError)
    retry.ts             ← withRetry<T> exponential backoff helper
  types/
    settings.ts          ← SettingsSchema interface
    export.ts            ← ExportJob, ExportResult types
```

---

## 2. Technology Stack

| Concern | Technology | Version |
|---------|-----------|---------|
| Language | TypeScript | 5.x, strict mode |
| UI framework | React | 18.x |
| Build | Vite | 5.x |
| Office integration | office-js | latest |
| Authentication | @azure/msal-browser | v3.x |
| Graph types | @microsoft/microsoft-graph-types | latest |
| Testing | Vitest (Jest-compatible API, Vite-native) | 2.x |
| Linting | ESLint + @typescript-eslint | latest |
| Manifest | JSON Unified Manifest | v2 |

---

## 3. Code Style Conventions

Conventions are aligned with the `llm-aggregator.ts` project to enable shared utility code and consistent readability across both projects.

### 3.1 TypeScript

- **Strict mode** on (`"strict": true` in tsconfig)
- **No `any`** in production code; use `unknown` + type narrowing
- **No semicolons** at end of statements
- **Single quotes** for string literals
- **ES2020 target**, ESNext modules
- `@/*` path alias for `src/*` imports
- Interface naming: `XData` for read shapes, `XCreateData` for create payloads, `XUpdateData` for partial update payloads

### 3.2 Utilities (shared pattern with llm-aggregator)

**`utils/logger.ts`** — conditional dev-only logging:
```typescript
export function debugLog(tag: string, ...args: unknown[]): void {
  if (process.env.NODE_ENV === 'development') {
    console.log(`[${tag}]`, ...args)
  }
}

export function debugError(tag: string, ...args: unknown[]): void {
  if (process.env.NODE_ENV === 'development') {
    console.error(`[${tag}]`, ...args)
  }
}
```

**`utils/retry.ts`** — exponential backoff (identical to llm-aggregator pattern):
```typescript
export async function withRetry<T>(
  fn: () => Promise<T>,
  maxRetries = 3,
  baseDelay = 500,
): Promise<T> {
  let attempt = 0
  while (true) {
    try {
      return await fn()
    } catch (error) {
      attempt++
      if (attempt >= maxRetries) throw error
      const delay = baseDelay * Math.pow(2, attempt - 1)
      debugLog('retry', `attempt ${attempt} in ${delay}ms...`)
      await new Promise(resolve => setTimeout(resolve, delay))
    }
  }
}
```

> These two utilities (`logger.ts`, `retry.ts`) are candidates for extraction into a shared package consumed by both Outlook2OneNote and llm-aggregator in the future.

### 3.3 Functions vs Classes

Prefer exported functions and interfaces over classes. Services are plain objects or modules with exported functions, not class instances — consistent with llm-aggregator's store/utility pattern.

---

## 4. Authentication Design

### 4.1 Strategy

`authService.ts` is the single source of truth for tokens. All other modules call `authService.getGraphToken()` — they never touch MSAL or Office.auth directly.

### 4.2 Account Type Policy

A compile-time / environment-level configuration controls which account types are accepted:

```typescript
// msalConfig.ts
type AccountTypePolicy = 'all' | 'msa-only' | 'entra-only'

const ACCOUNT_TYPE_POLICY: AccountTypePolicy =
  (process.env.ACCOUNT_TYPE_POLICY as AccountTypePolicy) ?? 'all'

// Authority mapping:
//   'all'         → 'https://login.microsoftonline.com/common'   (MSA + Entra)
//   'msa-only'    → 'https://login.microsoftonline.com/consumers'
//   'entra-only'  → 'https://login.microsoftonline.com/organizations'
```

This policy is set at build/deployment time by the IT administrator. It is not exposed in the user settings dialog. Entra Conditional Access policies are enforced by Microsoft's identity platform automatically — the add-in does not bypass them.

### 4.3 Token Acquisition Flow

```typescript
async getGraphToken(): Promise<string>
```

1. Check `sessionStorage` for a valid (non-expired) token → return immediately if found.
2. Try `Office.auth.getAccessToken({ allowSignInPrompt: false })` to get an SSO bootstrap token.
3. On SSO success: exchange via OBO flow (or use directly if scopes match — see OBO note in TASKS.md T-OBO).
4. On SSO failure (error codes 13000–13999): fall back to `msalInstance.acquireTokenSilent()`.
5. On silent failure: `msalInstance.acquireTokenPopup()`.
6. Store result in `sessionStorage` with expiry.
7. Throw `AuthError` with a user-readable message on all failures.

### 4.4 MSAL Configuration

```typescript
{
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: authorityForPolicy(ACCOUNT_TYPE_POLICY),
    redirectUri: `${window.location.origin}/auth/callback.html`,
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false,
  }
}
```

### 4.5 Scopes

See [docs/SPEC.md](SPEC.md) Section 5 for the required Graph API scopes and rationale.

---

## 5. Service Layer Contracts

### 5.1 graphClient.ts

```typescript
interface GraphClient {
  get<T>(path: string, params?: Record<string, string>): Promise<T>
  post<T>(path: string, body: unknown, contentType?: string): Promise<T>
}
```

- Attaches `Authorization: Bearer <token>` via `authService.getGraphToken()`.
- On 401: one token refresh attempt, retry once, then throw.
- On 429: waits `Retry-After` seconds (via `withRetry`), retries once.

### 5.2 mailService.ts

```typescript
interface MailService {
  getConversationMessages(conversationId: string): Promise<EmailMessage[]>
}
```

- Fetches all pages of `/me/messages?$filter=conversationId eq '{id}'&$select=...&$top=50`.
- Returns messages sorted by `receivedDateTime` ascending.
- Throws `ThreadTooLargeError` if message count exceeds 50 before returning.

### 5.3 settingsService.ts

```typescript
interface SettingsSchema {
  notebookId:           string | null
  notebookDisplayName:  string | null
  sectionNameFormat:    string       // default: '{subject} ({date})'
  sortOrder:            'asc' | 'desc'  // default: 'asc'
  includeAttachments:   boolean      // default: true
  preferredLink:        'web' | 'desktop' | 'both'  // default: 'both'
}

interface SettingsService {
  get(): SettingsSchema
  set(partial: Partial<SettingsSchema>): void
  save(): Promise<void>              // calls roamingSettings.saveAsync
}
```

---

## 6. OneNote Reusable Library

### 6.1 Boundary Rule

Everything under `src/onenote/` must have **zero imports from** `src/taskpane/`, `src/commands/`, `src/auth/`, `src/services/`, or any Office.js API. The only allowed dependencies are:

- `src/utils/` (logger, retry)
- `@microsoft/microsoft-graph-types`
- Standard TypeScript / DOM types

This makes `src/onenote/` independently portable.

### 6.2 Core Types (`onenote/types.ts`)

```typescript
// Generic email message — no Outlook/Graph-specific types
export interface EmailMessage {
  id: string
  subject: string
  from: EmailAddress
  toRecipients: EmailAddress[]
  ccRecipients: EmailAddress[]
  receivedDateTime: string          // ISO 8601
  bodyHtml: string                  // raw HTML body
  attachments: AttachmentMetadata[]
}

export interface EmailAddress {
  name: string
  address: string
}

export interface AttachmentMetadata {
  name: string
  size: number
  contentType: string
}

export interface Notebook { id: string; displayName: string; webUrl: string }
export interface Section  { id: string; displayName: string }
export interface Page     { id: string; webUrl: string; oneNoteClientUrl: string }
```

### 6.3 oneNoteService.ts

```typescript
interface OneNoteService {
  listNotebooks(): Promise<Notebook[]>
  createSection(notebookId: string, sectionName: string): Promise<Section>
  createPage(sectionId: string, message: EmailMessage, title: string): Promise<Page>
}
```

The service accepts a `GraphClient` instance injected at construction — it does not create its own HTTP client. This allows the same service to be wired to different Graph client implementations in different host projects.

### 6.4 pageBuilder.ts

Accepts `EmailMessage` (the generic type, not Graph's `Message`) and produces OneNote-compatible HTML. This is the key abstraction: any project that can produce an `EmailMessage` can use the page builder without knowing about Outlook or Graph Mail.

---

## 7. Export Orchestration

The export pipeline runs in `ExportView.tsx` (delegating to an `exportOrchestrator.ts` helper):

```
1. Validate: notebook selected? → if not, show inline prompt to open settings
2. Read conversationId from Office.context.mailbox.item
3. mailService.getConversationMessages(conversationId)
   → throws ThreadTooLargeError if > 50 messages → show advisory message
   → emit progress: 'Fetching messages…'
4. Apply sort order from settings
5. Build section name from settings.sectionNameFormat
6. oneNoteService.createSection(notebookId, sectionName)
   → emit progress: 'Creating OneNote section…'
7. For each message (index i of n):
   a. oneNoteService.createPage(sectionId, emailMessage, title)
   → emit progress: `Exporting message ${i + 1} of ${n}…`
8. Return page links filtered by settings.preferredLink
   → display success with link(s)
```

All steps wrapped in try/catch; errors surface via `ErrorBanner` with retry.

---

## 8. Settings UI

Settings are accessed exclusively via the **⚙ gear icon** in `AppHeader.tsx`. There is no second ribbon button.

Interaction model (matching llm-aggregator `SettingsDialog` pattern):
- Gear icon click → `showSettings` state set to true → `SettingsView` renders as overlay
- Click outside overlay → dismiss without saving
- Escape key → dismiss without saving
- Ctrl+Enter or Save button → `settingsService.set(...)` + `settingsService.save()` → dismiss
- Settings are not persisted until Save is explicitly triggered

Admin-controlled settings (account type policy) are read-only environment values and are **not rendered** in the settings dialog.

---

## 9. Thread Size Limit

- Hard cap: **50 messages**
- `mailService` fetches up to 51 messages (one over limit) to detect the breach without loading the full thread
- On breach: `ThreadTooLargeError` thrown with `count` property
- `ExportView` catches this and renders: _"This thread has {N}+ messages, which exceeds the 50-message limit. Please archive or delete older messages in Outlook to reduce the thread size, then try again."_
- No partial export is attempted when the limit is exceeded

---

## 10. Testing Strategy

| Layer | What to test | Tool |
|-------|-------------|------|
| Unit | `pageBuilder`, `htmlSanitizer`, `settingsService`, `dateFormatter` | Vitest |
| Unit | `authService` token logic (MSAL mocked) | Vitest |
| Unit | `mailService`, `oneNoteService` (graphClient mocked) | Vitest |
| Integration | Export orchestration end-to-end (all services mocked at fetch boundary) | Vitest |
| Manual | Full test plan — see [docs/TEST-PLAN.md](TEST-PLAN.md) | Human |
| E2E automated | Playwright | Deferred (future release) |

---

## 11. llm-aggregator Integration

### 11.1 Vision

`llm-aggregator` is a standalone Electron + Vue app that aggregates QA pairs and threads from multiple sources. The planned pipeline is:

```
Outlook email thread
       │
       ▼
Outlook2OneNote (add-in)
  mailService.getConversationMessages()
       │  EmailMessage[]
       ▼
  [future] llm-aggregator import endpoint
       │  ThreadData + QAPairData
       ▼
  [future] llm-aggregator → oneNoteService.createPage()
       │
       ▼
  OneNote page
```

In this model, llm-aggregator becomes the **unified QA/thread importer** and uses the `src/onenote/` library directly to export to OneNote. Outlook2OneNote's role is to produce `EmailMessage[]` and pass them to the aggregator.

### 11.2 Shared Code Candidates

| Module | Status |
|--------|--------|
| `utils/retry.ts` | Identical pattern already in llm-aggregator — extract to shared package |
| `utils/logger.ts` | Identical pattern already in llm-aggregator — extract to shared package |
| `onenote/types.ts` | New — consumed by both projects once integration is done |
| `onenote/oneNoteService.ts` | New — consumed by llm-aggregator in future |
| `onenote/pageBuilder.ts` | New — consumed by llm-aggregator in future |

### 11.3 v2 Constraints

- v2 of Outlook2OneNote exports directly to OneNote (not via aggregator) — the aggregator pipeline is future scope
- The `src/onenote/` boundary rule (Section 6.1) must be enforced from the start so the future integration requires no refactoring
- `EmailMessage` type (Section 6.2) is the contract between the two projects; changes to it are breaking changes

---

## 12. Build & Dev Setup

```
npm install
npm run dev        # webpack-dev-server with HTTPS (required for Office add-ins)
npm run build      # production bundle
npm run test       # Jest
npm run lint       # ESLint
```

Environment variables (injected via webpack `DefinePlugin` from `.env`):

| Variable | Values | Default |
|----------|--------|---------|
| `CLIENT_ID` | Azure app registration GUID | required |
| `TENANT_ID` | `common` / specific tenant GUID | `common` |
| `ACCOUNT_TYPE_POLICY` | `all` / `msa-only` / `entra-only` | `all` |

---

*Last updated: 2026-02-28*
