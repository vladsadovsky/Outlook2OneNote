# Claude Code Project Instructions — Outlook2OneNote

## Project Intent

This is a **ground-up rewrite** (v2) of an Outlook add-in that exports email threads to OneNote.
The `master` branch contains a partially working v1 — **use it for reference only** (feature set, auth patterns, API calls).
All new work goes on `v2/main` and feature branches off it.

## Model Preferences

- **Planning, architecture, spec review** → switch to Opus: `/model opus`
- **Coding, refactoring, implementation** → Sonnet (default)
- When in doubt about an architectural decision, spawn an Opus planning session before writing code.

## Workflow: Spec-First

No code is written without a corresponding spec item. Follow this sequence:

```
docs/SPEC.md     ← functional requirements + acceptance criteria  (write first)
docs/DESIGN.md   ← module architecture, contracts, data flow      (write second)
docs/TASKS.md    ← trackable work items with status               (drive implementation)
src/             ← implementation driven by spec
tests/           ← acceptance tests tied to spec items
```

Use `TodoWrite` for in-session progress tracking. Update `docs/TASKS.md` to persist across sessions.

## Agent Safety Rules (all coding agents)

These rules are mandatory for any agent editing this repo.

- **Never erase existing note/log content unless explicitly asked.**
- For `DEV_NOTES.md`, `docs/DEV-NOTES.md`, `docs/TASK-LOG.md`, and `docs/TASKS.md`, use **append-or-merge only by default**.
- If the user explicitly asks to edit/rewrite specific sections, do it, but preserve unrelated dated entries unless removal is explicitly requested.
- Before editing any documentation file, read current contents and preserve existing dated entries.
- If a requested change could replace or conflict with existing content, ask for confirmation before destructive edits.
- After documentation edits, summarize **added vs changed vs removed** sections.
- If two files appear to serve a similar purpose, do not consolidate/delete without explicit user instruction.

## Branch Strategy

```
master           ← v1 reference, do not modify
v2/main          ← clean rewrite base
v2/feat/<name>   ← feature branches off v2/main
```

Create `v2/main` from `master` HEAD before starting any implementation.

## Technology Stack (Recommended for v2)

- **TypeScript** — Office.js, MSAL, and Graph API all have full type coverage; catches auth/API bugs early
- **Webpack** — already configured in v1, carry forward
- **Jest** — unit/integration testing
- **Office Add-in Unified Manifest (JSON)** — new format, preferred over legacy XML unless AppSource compatibility with classic Outlook forces XML
- **MSAL Browser v3** — upgrade from v2 used in v1
- **@microsoft/microsoft-graph-types** — type Graph API responses

## Key Technical Context

### Authentication (the hard problem)
The core difficulty in v1 was auth. The target behavior:
- If user is authenticated with Office (SSO), use those credentials silently
- If not, show a single popup (not a redirect) — matching Microsoft's own "Save to OneNote" add-in UX
- No client secrets in frontend code
- Token storage in sessionStorage only (cleared on tab close)
- Pattern: `Office.auth.getAccessToken()` first → on failure → MSAL popup fallback

### Platform Targets
- Outlook Web (primary)
- Outlook Desktop New / WebView2 (secondary)
- Outlook Classic (TBD — see pending questions below)

### Core Export Flow
1. Read selected email's `conversationId` from `Office.context.mailbox.item`
2. Fetch full thread via Graph API (`/me/messages?$filter=conversationId eq '...'`)
3. Create new OneNote section in user-selected notebook
4. Create one page per email, sorted chronologically

### Graph API Scopes Required
- `Mail.Read`
- `Notes.ReadWrite`
- `offline_access` (for refresh tokens)

---

## PENDING: Clarifying Questions for Spec

**Before writing docs/SPEC.md, get answers to these questions:**

### Authentication & Accounts
1. Personal Microsoft accounts only, or also work/school accounts (Entra ID)?
2. "Consistent with native Microsoft" — visual only (popup, not redirect), or technically identical token scopes to built-in Save to OneNote?

### Platform Scope
3. Outlook Classic support — must-have for v1, or defer to v2+?
4. OneNote target — classic desktop, OneNote for Web, or both?

### Settings / Preferences
5. Essential v1 settings list? Current known: notebook selection (persisted via roamingSettings). Others?
   - Section naming format (customizable)?
   - Email sort order (configurable)?
   - Include attachments?

### Email Content
6. HTML email body fidelity in OneNote — exact HTML preserved, or plain text acceptable for v1?
7. Attachments — out of scope for v1, or include at least links/metadata?

### Distribution
8. Personal use only, or targeting Microsoft AppSource eventually?
   (Affects manifest format: XML vs JSON Unified Manifest)

### Stack
9. TypeScript — yes, or prefer to stay with JavaScript?
10. Testing framework preference, or no preference?

---

## Reference: v1 File Structure (for learning, not porting)

```
src/
  taskpane/
    taskpane.js          ← main UI logic
    taskpane.html
    taskpanemethods.js   ← export orchestration
    email-service.js     ← Graph API email calls
    onenote-service.js   ← Graph API OneNote calls
  auth/
    callback.html        ← MSAL redirect callback
  commands/
    commands.js          ← ribbon button handler
  common/
    app-state.js         ← shared state (likely)
auth/
  callback.html          ← outer MSAL callback
manifest.xml             ← add-in manifest (XML format)
webpack.config.js
```

---

*Created: 2026-02-28 — context carried forward from planning session*
*Next step: answer pending questions above, then write docs/SPEC.md*
