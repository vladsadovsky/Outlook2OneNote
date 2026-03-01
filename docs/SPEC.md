# Outlook2OneNote v2 — Functional Specification

**Version:** 2.1
**Date:** 2026-02-28
**Status:** Draft

---

## 1. Overview

Outlook2OneNote is a Microsoft Outlook add-in that exports an email conversation thread to a new section in Microsoft OneNote. The user selects a notebook, triggers the export from a ribbon button, and receives a structured OneNote section with one page per email, sorted chronologically.

The OneNote integration layer is architected for reuse: it is a self-contained module with no dependencies on Outlook or the add-in host, so it can be consumed by other projects (see Section 10).

---

## 2. Scope

### 2.1 In Scope (v2)

- Export a full email conversation thread to OneNote
- Support both personal Microsoft accounts (MSA) and work/school accounts (Entra ID)
- Account type can be restricted per deployment (see Section 3, US-01)
- Target platforms: Outlook Web App, Outlook Desktop New (WebView2)
- OneNote target: OneNote for Web + Desktop (via Graph API — pages sync automatically); target platform can be restricted in settings
- TypeScript implementation
- Distribution: sideloaded initially; built for AppSource submission (JSON Unified Manifest)
- Comprehensive manual test plan covering functional, reliability, performance, and stability
- Automated testing (Playwright) deferred to a future release

### 2.2 Out of Scope (v2)

- Outlook Classic (COM add-in) — deferred to future version
- Full attachment binary download/embed — metadata only in v2
- Shared mailbox support
- Mobile Outlook
- Playwright / automated end-to-end tests (future release)

---

## 3. User Stories & Acceptance Criteria

### US-01 — Authenticate
**As a user**, I want to sign in with my Microsoft account once, so that subsequent exports work without re-authentication.

**Acceptance Criteria:**
- AC-01.1: If Office SSO is available (`Office.auth.getAccessToken()` succeeds), the add-in acquires a Graph token silently — no UI shown.
- AC-01.2: If SSO fails, a single MSAL popup appears for interactive login — no redirect, no embedded auth page.
- AC-01.3: Both MSA (personal) and Entra ID (work/school) accounts are accepted by default.
- AC-01.4: A per-deployment setting (environment variable / manifest configuration) can restrict accepted account types to MSA-only or Entra-only. This setting is not modifiable by the end user — it is an admin/IT control.
- AC-01.5: The add-in respects Microsoft Entra Conditional Access policies and Group Policy restrictions: if an IT administrator has configured tenant policies that disallow personal accounts, only Entra accounts are accepted and the add-in does not attempt MSA login.
- AC-01.6: Tokens are stored in `sessionStorage` only; cleared when the browser tab is closed.
- AC-01.7: No client secrets are present in front-end code.
- AC-01.8: On token expiry the add-in attempts silent refresh before prompting the user.

### US-02 — Select Notebook
**As a user**, I want to choose which OneNote notebook my threads are exported to, so that I can keep work and personal notes separate.

**Acceptance Criteria:**
- AC-02.1: The task pane lists all notebooks the user has access to via Graph API.
- AC-02.2: The selected notebook is persisted via `Office.context.roamingSettings` and pre-selected on subsequent opens.
- AC-02.3: If the persisted notebook no longer exists, the add-in prompts re-selection.
- AC-02.4: The notebook list refreshes on demand (refresh button).
- AC-02.5: A platform restriction setting (admin/IT-controlled, or user-configurable) can limit notebook access to OneNote for Web only, or require both Web and Desktop sync. This does not change the Graph API call — it only affects which `webUrl` vs `oneNoteClientUrl` link is shown to the user after export.

### US-03 — Export Email Thread
**As a user**, I want to export the currently selected email conversation to OneNote with one click, so I have a permanent, readable record.

**Acceptance Criteria:**
- AC-03.1: Triggering export reads `conversationId` from `Office.context.mailbox.item`.
- AC-03.2: All messages in the conversation are fetched via Graph API (`/me/messages?$filter=conversationId eq '...'`), including messages across folders (Sent Items, Inbox, etc.).
- AC-03.3: If the conversation contains more than **50 messages**, the export is blocked. The user is shown a message: _"This thread has N messages, which exceeds the 50-message limit. Please archive or delete older messages in Outlook to reduce the thread size, then try again."_
- AC-03.4: Messages are sorted chronologically (oldest first) by `receivedDateTime`, unless the user has configured reverse order in settings.
- AC-03.5: A new OneNote section is created in the selected notebook named per the configured format (default: email subject + date).
- AC-03.6: Each message is created as a separate OneNote page within that section.
- AC-03.7: Page title format: `[receivedDateTime ISO date] — [sender display name]`.
- AC-03.8: Page body contains: sender, recipients (To/CC), date/time, subject, and message body.
- AC-03.9: Message body is rendered as best-effort HTML — structure and formatting preserved; complex CSS and external images may be stripped.
- AC-03.10: Attachment metadata (filename, size, MIME type) is listed at the bottom of each page when the attachments setting is enabled; no binary download.
- AC-03.11: Progress is shown to the user (e.g. "Exporting 5 of 12 messages…").
- AC-03.12: On success, the user is shown a link to the created OneNote section. If the platform setting allows Web access, a web link is shown; if Desktop access is also enabled, both links are shown.
- AC-03.13: On failure, a meaningful error message is shown with a retry option.

### US-04 — Configure Settings
**As a user**, I want to configure export preferences so that the output matches my workflow.

**Acceptance Criteria:**
- AC-04.1: Settings are accessible via a **gear icon (⚙) in the task pane header** — no additional ribbon button.
- AC-04.2: Clicking the gear icon opens a settings overlay/dialog within the task pane (not a new window or browser tab).
- AC-04.3: Pressing Escape closes the settings dialog without saving. Ctrl+Enter (or a Save button) saves and closes.
- AC-04.4: **Notebook selection** — as per US-02.
- AC-04.5: **Section naming format** — user can configure a template, e.g. `{subject} ({date})` or `{date} — {subject}`. Default: `{subject} ({date})`.
- AC-04.6: **Email sort order** — chronological (oldest first) or reverse-chronological (newest first). Default: chronological.
- AC-04.7: **Include attachment metadata** — toggle on/off. Default: on.
- AC-04.8: **Preferred OneNote link** — Web only, Desktop only, or Both. Default: Both. (Only affects which link is shown after export — does not restrict notebook access.)
- AC-04.9: All user-facing settings are persisted via `Office.context.roamingSettings`.
- AC-04.10: Admin-controlled settings (account type restriction, IT policy) are not shown in the user settings dialog.

### US-05 — Ribbon Button
**As a user**, I want a clearly labelled ribbon button in Outlook to open the export task pane.

**Acceptance Criteria:**
- AC-05.1: A single "Export to OneNote" button appears in the Outlook ribbon (Home tab, Message tab).
- AC-05.2: The button opens the add-in task pane.
- AC-05.3: The button is visible in Outlook Web and Outlook New Desktop.
- AC-05.4: No additional ribbon buttons are added (settings, help, etc. are accessible from within the task pane).

---

## 4. Non-Functional Requirements

| ID | Requirement |
|----|-------------|
| NFR-01 | Task pane initial load ≤ 3 seconds on a standard broadband connection |
| NFR-02 | Export of a 20-message thread completes in ≤ 30 seconds |
| NFR-03 | No client secrets or tokens are written to localStorage or cookies |
| NFR-04 | Add-in passes Office Add-in Validator checks (for AppSource readiness) |
| NFR-05 | All Graph API calls use the least-privilege scopes required |
| NFR-06 | TypeScript strict mode enabled; no `any` types in production code |
| NFR-07 | Thread size hard cap: 50 messages (AC-03.3) |

---

## 5. Graph API Scopes

| Scope | Reason |
|-------|--------|
| `Mail.Read` | Read email messages and conversation threads |
| `Notes.ReadWrite` | Create notebooks, sections, and pages in OneNote |
| `offline_access` | Obtain refresh tokens for silent re-authentication |
| `User.Read` | Read basic profile (display name, email) |

---

## 6. Auth Flow

SSO is attempted first; on failure, silent MSAL refresh; on failure, interactive popup. Account type policy filters which accounts the popup accepts. Tokens are stored in `sessionStorage` only.

For the full token acquisition sequence and MSAL configuration, see [docs/DESIGN.md](DESIGN.md) Section 4.

---

## 7. Data Flow Summary

```
Outlook item
  └─ conversationId
        │
        ▼
Graph /me/messages (filter by conversationId)
  └─ array of Message objects (max 50; block if exceeded)
        │
        ▼
Sort by receivedDateTime (per settings)
        │
        ▼
Graph POST /notebooks/{id}/sections  (create section)
        │
        ▼
For each message:
  Graph POST /sections/{id}/pages    (create page with OOXML/HTML body)
        │
        ▼
Return section webUrl / oneNoteClientUrl → show link(s) to user
```

---

## 8. Manifest

- Format: **JSON Unified Manifest** (Office Add-ins v2)
- Add-in type: Task pane + Commands (single ribbon button)
- Requirements: `Mailbox 1.10` minimum
- Permissions: `ReadItem` (read selected message)

---

## 9. Testing

The full manual test plan (functional, reliability, performance, stability) is in [docs/TEST-PLAN.md](TEST-PLAN.md). Automated Jest unit/integration tests are part of the implementation. Playwright E2E is deferred — see Section 11.

---

## 10. OneNote Layer — Reuse Architecture

The `oneNoteService` module and its dependent types (`Notebook`, `Section`, `Page`) are structured as a standalone, self-contained library with no imports from Outlook-specific code. This enables:

- **Direct reuse** in other TypeScript projects that need OneNote export (e.g. llm-aggregator)
- **Pipeline architecture**: Outlook2OneNote → (future) llm-aggregator → OneNote, where the aggregator acts as a unified QA/thread importer that uses the same OneNote library

See [docs/DESIGN.md](DESIGN.md) Section 11 for the reuse boundary, shared types, and the planned llm-aggregator integration pipeline.

---

## 11. Out-of-Scope / Future Considerations

- Full attachment download and embedding in OneNote
- Outlook Classic / COM add-in
- Shared mailbox / delegate access
- Mobile Outlook
- Bulk export (multiple conversations at once)
- Two-way sync (edit in OneNote, reflect back to email)
- Playwright automated test suite
- llm-aggregator as export intermediary (full v2 integration)

---

*Last updated: 2026-02-28*
