# Outlook2OneNote v2 — Manual Test Plan

**Version:** 1.0
**Date:** 2026-02-28
**Spec ref:** docs/SPEC.md
**Tester:** Human (automated E2E deferred — see TASKS.md T-F01)

---

## How to Use This Plan

- Work through sections in order for a full regression pass.
- Each test case has: **ID**, **preconditions**, numbered **steps**, and **expected result**.
- Mark results: ✅ Pass · ❌ Fail · ⚠ Partial · ⏭ Skipped (note reason)
- Record failures with: actual result, screenshot/log if available, and link to bug.
- "Clean state" means: clear browser sessionStorage, clear roamingSettings (use the reset helper in Settings), and reload the add-in.

---

## Test Environments

Run the full plan in each environment before a release:

| Env | Browser / Host | Account type |
|-----|---------------|-------------|
| OWA-MSA | Outlook Web App (Edge) | Personal MSA |
| OWA-Entra | Outlook Web App (Edge) | Work/school Entra ID |
| NEW-DT | Outlook Desktop New (WebView2) | Entra ID |

---

## Section 1 — Authentication

### TC-AUTH-01: SSO silent sign-in (Entra ID)
**Preconditions:** User is already signed into Outlook Web with an Entra ID account. Clean state.
1. Open Outlook Web, navigate to any email.
2. Click the "Export to OneNote" ribbon button.
3. Observe the task pane loading sequence.

**Expected:** Task pane loads without any login prompt or popup. User is considered authenticated. No `sessionStorage` auth error keys present.

---

### TC-AUTH-02: SSO silent sign-in (MSA)
**Preconditions:** User is signed into Outlook Web with a personal MSA account. Clean state.
1. Open Outlook Web, navigate to any email.
2. Click the "Export to OneNote" ribbon button.

**Expected:** Task pane loads without any login prompt. Silent SSO succeeds.

---

### TC-AUTH-03: MSAL popup fallback — Entra ID
**Preconditions:** SSO is unavailable (sideloaded add-in on a fresh profile, or SSO intentionally blocked via `ACCOUNT_TYPE_POLICY=entra-only` with mismatched session). Clean state.
1. Open the task pane.
2. Observe: a popup window appears.
3. Complete sign-in with an Entra ID account in the popup.

**Expected:** Popup closes automatically. Task pane proceeds to the notebook selection state. No redirect occurs — popup only.

---

### TC-AUTH-04: MSAL popup fallback — MSA
**Preconditions:** Same as TC-AUTH-03, using a personal MSA account.
1. Open the task pane.
2. Complete sign-in with a personal MSA account in the popup.

**Expected:** Same as TC-AUTH-03.

---

### TC-AUTH-05: Token persists within session
**Preconditions:** Complete TC-AUTH-01 or TC-AUTH-03 successfully.
1. Close the task pane.
2. Navigate to a different email.
3. Re-open the task pane.

**Expected:** Task pane opens directly to the export state — no re-authentication prompt. Token retrieved from `sessionStorage`.

---

### TC-AUTH-06: Token cleared on tab close
**Preconditions:** Complete any auth test. Confirm token in `sessionStorage`.
1. Close the Outlook browser tab entirely.
2. Open a new tab and navigate back to Outlook Web.
3. Open the task pane.

**Expected:** Authentication is required again (SSO re-attempted first, then popup if needed). Token was not persisted to `localStorage` or cookies.

---

### TC-AUTH-07: Account type policy — `entra-only` blocks MSA
**Preconditions:** Add-in built with `ACCOUNT_TYPE_POLICY=entra-only`. Clean state.
1. Open the task pane.
2. In the MSAL popup, attempt to sign in with a personal MSA account (@outlook.com, @hotmail.com).

**Expected:** Sign-in fails or is rejected. An appropriate error message is shown. No token is stored.

---

### TC-AUTH-08: Account type policy — `msa-only` blocks Entra
**Preconditions:** Add-in built with `ACCOUNT_TYPE_POLICY=msa-only`. Clean state.
1. Open the task pane.
2. In the MSAL popup, attempt to sign in with a work/school Entra ID account.

**Expected:** Sign-in fails or is rejected. An appropriate error message is shown.

---

### TC-AUTH-09: Token expiry — silent refresh
**Preconditions:** Authenticated session active. Manually expire the token by editing `sessionStorage` (set `exp` to a past timestamp) or wait for natural expiry.
1. Trigger an export action.

**Expected:** Add-in silently refreshes the token without prompting the user. Export proceeds normally.

---

### TC-AUTH-10: Token expiry — refresh fails → popup
**Preconditions:** Authenticated session active. Expire token AND simulate silent refresh failure (e.g. revoke refresh token in Azure portal).
1. Trigger an export action.

**Expected:** Add-in shows a sign-in prompt (popup). After re-authentication, the action resumes or the user is asked to retry.

---

## Section 2 — Settings

### TC-SETTINGS-01: Open settings via gear icon
**Preconditions:** Task pane open, authenticated.
1. Locate the ⚙ gear icon in the task pane header.
2. Click it.

**Expected:** Settings overlay appears. Main export view is still visible behind a semi-transparent backdrop.

---

### TC-SETTINGS-02: Close settings with Escape
**Preconditions:** Settings overlay open. Change at least one value but do not save.
1. Press Escape.

**Expected:** Overlay closes. Changed values are discarded. Original settings are unchanged.

---

### TC-SETTINGS-03: Close settings by clicking outside
**Preconditions:** Settings overlay open.
1. Click the semi-transparent backdrop area outside the dialog.

**Expected:** Overlay closes. Changes discarded.

---

### TC-SETTINGS-04: Save settings with Ctrl+Enter
**Preconditions:** Settings overlay open.
1. Change section naming format to `{date} — {subject}`.
2. Press Ctrl+Enter.

**Expected:** Overlay closes. Changed value is saved. Re-opening settings confirms the new value is persisted.

---

### TC-SETTINGS-05: Save settings with Save button
**Preconditions:** Settings overlay open.
1. Change sort order to reverse-chronological.
2. Click Save.

**Expected:** Overlay closes. Value persisted to roamingSettings. Confirmed on re-open.

---

### TC-SETTINGS-06: Settings persist across sessions
**Preconditions:** Save a non-default value in settings (e.g. reverse sort order).
1. Close the task pane.
2. Close and reopen the browser tab.
3. Open the task pane and navigate to settings.

**Expected:** The previously saved non-default value is still present.

---

### TC-SETTINGS-07: Settings persist across devices
**Preconditions:** Save settings on Device A. Log in to the same Microsoft account on Device B.
1. Open Outlook Web on Device B.
2. Open the task pane and navigate to settings.

**Expected:** The settings from Device A are reflected (roamingSettings sync).

---

### TC-SETTINGS-08: No second ribbon button
**Preconditions:** Add-in installed/sideloaded.
1. Examine the Outlook ribbon (Home tab, Message tab).

**Expected:** Exactly one add-in button: "Export to OneNote". No settings button, no help button in the ribbon.

---

### TC-SETTINGS-09: Admin account type policy not shown in settings
**Preconditions:** Add-in built with any `ACCOUNT_TYPE_POLICY` value.
1. Open settings overlay.

**Expected:** There is no field for account type restriction visible to the user.

---

## Section 3 — Notebook Selection

### TC-NB-01: Notebook list loads
**Preconditions:** Authenticated. No notebook previously selected.
1. Open the task pane (or navigate to notebook picker).

**Expected:** A list of the user's OneNote notebooks is displayed. List matches notebooks visible in OneNote for Web.

---

### TC-NB-02: Notebook selection persists
**Preconditions:** TC-NB-01 complete.
1. Select a notebook from the list.
2. Close and reopen the task pane.

**Expected:** The previously selected notebook is pre-selected.

---

### TC-NB-03: Notebook list refreshes
**Preconditions:** Authenticated.
1. Create a new notebook directly in OneNote for Web.
2. Return to the add-in task pane.
3. Click the refresh button on the notebook picker.

**Expected:** The newly created notebook appears in the list.

---

### TC-NB-04: Missing notebook detection
**Preconditions:** A notebook is selected and persisted. Delete that notebook from OneNote.
1. Open the task pane.

**Expected:** Add-in detects the notebook no longer exists and shows a prompt asking the user to re-select. Does not silently continue with a broken notebook ID.

---

### TC-NB-05: No notebook selected — export blocked
**Preconditions:** Clean state, no notebook selected.
1. Select an email with a conversation thread.
2. Click the Export button.

**Expected:** Export does not start. An inline message prompts the user to select a notebook in settings first.

---

## Section 4 — Export — Core Flow

### TC-EXP-01: Basic export — small thread
**Preconditions:** Authenticated, notebook selected. Select an email that is part of a thread with 3–5 messages.
1. Click Export.
2. Observe progress indicator.
3. Wait for completion.

**Expected:**
- Progress updates: "Fetching messages…", then "Exporting message N of M…" for each message.
- A new section appears in the selected OneNote notebook named per the configured format.
- The section contains one page per message.
- Pages are ordered chronologically (oldest first with default settings).

---

### TC-EXP-02: Page content — headers
**Preconditions:** TC-EXP-01 complete.
1. Open one of the created OneNote pages.

**Expected:** Page contains: From, To, CC, Date, Subject fields. All populated correctly and matching the original email.

---

### TC-EXP-03: Page content — body HTML
**Preconditions:** Source email contains bold text, bullet lists, and inline images.
1. Export that thread.
2. Inspect the page body in OneNote.

**Expected:** Bold and bullet list formatting is preserved. Inline images from external URLs may be absent (stripped). No `<script>` or `<iframe>` tags present. No JavaScript event attributes (e.g. `onclick`).

---

### TC-EXP-04: Page content — attachment metadata
**Preconditions:** Source email has file attachments. `includeAttachments` setting is on.
1. Export that thread.
2. Inspect the page in OneNote.

**Expected:** An "Attachments" section at the bottom of the page lists each attachment's filename, size, and MIME type. No actual file is embedded.

---

### TC-EXP-05: Attachment metadata suppressed when disabled
**Preconditions:** Set `includeAttachments` to off in settings. Source email has attachments.
1. Export the thread.
2. Inspect the page.

**Expected:** No "Attachments" section on the page.

---

### TC-EXP-06: Sort order — chronological (default)
**Preconditions:** Thread with at least 3 messages. Default (ascending) sort order.
1. Export and inspect pages in OneNote.

**Expected:** Page 1 is the oldest message, last page is the newest.

---

### TC-EXP-07: Sort order — reverse-chronological
**Preconditions:** Thread with at least 3 messages. Sort order set to descending in settings.
1. Export and inspect pages.

**Expected:** Page 1 is the newest message, last page is the oldest.

---

### TC-EXP-08: Section naming — default format
**Preconditions:** Default section naming format `{subject} ({date})`.
1. Export a thread with subject "Project Alpha Review".
2. Check the section name in OneNote.

**Expected:** Section named e.g. `Project Alpha Review (2026-02-28)`.

---

### TC-EXP-09: Section naming — custom format
**Preconditions:** Section naming format changed to `{date} — {subject}`.
1. Export the same thread.

**Expected:** Section named e.g. `2026-02-28 — Project Alpha Review`.

---

### TC-EXP-10: Success state — web link shown
**Preconditions:** `preferredLink` = Both or Web.
1. Export completes successfully.

**Expected:** A clickable link to the OneNote section (web URL) is shown in the task pane.

---

### TC-EXP-11: Success state — desktop link shown
**Preconditions:** `preferredLink` = Both or Desktop.
1. Export completes successfully.

**Expected:** A clickable link that opens the section in the desktop OneNote app is shown.

---

### TC-EXP-12: Messages fetched across folders
**Preconditions:** A thread where some messages are in Inbox and some are in Sent Items.
1. Export the thread.

**Expected:** Pages in OneNote include both received and sent messages from the thread.

---

## Section 5 — Thread Size Limit

### TC-LIMIT-01: Thread exactly at limit (50 messages)
**Preconditions:** A thread with exactly 50 messages.
1. Export the thread.

**Expected:** Export proceeds normally. All 50 pages created in OneNote.

---

### TC-LIMIT-02: Thread exceeds limit (51+ messages)
**Preconditions:** A thread with 51 or more messages.
1. Attempt to export the thread.

**Expected:**
- Export does not start.
- Message shown: _"This thread has N+ messages, which exceeds the 50-message limit. Please archive or delete older messages in Outlook to reduce the thread size, then try again."_
- No partial section or pages are created in OneNote.

---

### TC-LIMIT-03: After trimming an oversized thread
**Preconditions:** TC-LIMIT-02 scenario. User deletes/archives messages to bring thread to ≤ 50.
1. Re-attempt the export.

**Expected:** Export succeeds normally.

---

## Section 6 — Reliability

### TC-REL-01: Network interruption during message fetch
**Preconditions:** Thread with 10+ messages. Use browser DevTools to throttle/offline network after export starts but before fetch completes.
1. Start export.
2. Go offline via DevTools.

**Expected:** Add-in shows a meaningful error (not a blank crash). Retry option is presented. Re-enabling network and retrying resumes the export.

---

### TC-REL-02: Network interruption during page creation
**Preconditions:** Thread with 5+ messages. Go offline after the first OneNote page is created.
1. Start export.
2. Go offline mid-page-creation.

**Expected:** Error is shown with retry option. On retry (with network restored), export either restarts cleanly or resumes. No duplicate section created on retry.

---

### TC-REL-03: Graph API 429 throttle response
**Preconditions:** Use a proxy or mock server to inject a `429 Retry-After: 5` response for one of the page creation calls.
1. Run export.

**Expected:** Add-in waits the specified `Retry-After` duration and retries automatically. Export completes successfully. User sees the progress indicator pause briefly, then continue.

---

### TC-REL-04: OneNote section creation failure
**Preconditions:** Use a proxy to return a `500` error on the `POST /sections` call.
1. Attempt export.

**Expected:** Error banner shown: section could not be created. Retry option available. No partial state left in OneNote.

---

### TC-REL-05: Single page creation failure mid-export
**Preconditions:** Use a proxy to return `500` on page 3 of 5 during export.
1. Attempt export.

**Expected:** Error banner shown, indicating failure at message N. Retry option available. Section was created; partial pages (1–2) may exist — this is acceptable and noted in the error message.

---

### TC-REL-06: Retry after error succeeds
**Preconditions:** TC-REL-01 or TC-REL-04 — error state reached.
1. Restore network / remove proxy fault.
2. Click Retry.

**Expected:** Export completes successfully.

---

## Section 7 — Performance

### TC-PERF-01: Task pane initial load time
**Preconditions:** Clear browser cache. Standard broadband connection (≥10 Mbps).
1. Click the ribbon button to open the task pane.
2. Measure time from click to task pane fully interactive (notebook picker visible, not loading spinner).

**Expected:** ≤ 3 seconds. (NFR-01)

**Measurement method:** Browser DevTools Network tab — note DOMContentLoaded + resource load times. Repeat 3 times and take average.

---

### TC-PERF-02: 20-message thread export time
**Preconditions:** A thread with exactly 20 messages. Standard broadband.
1. Start export and note start time.
2. Note time when success message appears.

**Expected:** ≤ 30 seconds total. (NFR-02)

**Measurement method:** Browser DevTools performance timeline, or stopwatch. Repeat 3 times.

---

### TC-PERF-03: Maximum thread (50 messages) export time
**Preconditions:** A thread with exactly 50 messages.
1. Export and time as above.

**Expected:** Completes within a reasonable time (document actual result; target guidance: ≤ 75 seconds). No timeout or crash.

---

### TC-PERF-04: Repeated exports — no memory leak / slowdown
**Preconditions:** Task pane remains open.
1. Export 5 different threads sequentially without closing the task pane.
2. Observe task pane responsiveness after each export.

**Expected:** No progressive slowdown. No visible memory growth in DevTools Memory tab between exports.

---

## Section 8 — Stability

### TC-STAB-01: Task pane survives Outlook page refresh
**Preconditions:** Task pane open, authenticated, notebook selected.
1. Refresh the Outlook Web browser tab (F5).
2. Re-open the task pane.

**Expected:** Task pane reloads cleanly. Settings (notebook selection) are restored from roamingSettings. Re-authentication may be required (SSO should handle silently).

---

### TC-STAB-02: Repeated open/close cycles
**Preconditions:** Task pane available.
1. Open and close the task pane 10 times in quick succession.

**Expected:** No crashes, blank panes, or stuck loading spinners. 10th open is identical to 1st.

---

### TC-STAB-03: Export while switching emails
**Preconditions:** Thread export in progress (progress indicator visible).
1. Click a different email in the Outlook message list.

**Expected:** Export continues to completion (it was already reading the conversationId at start). Export does not silently switch to the newly selected email's conversation.

---

### TC-STAB-04: Multiple rapid export clicks
**Preconditions:** Authenticated, notebook selected, email with thread selected.
1. Click the Export button 3 times rapidly.

**Expected:** Only one export runs. Subsequent clicks are ignored or debounced while an export is in progress.

---

### TC-STAB-05: Session timeout / Office re-auth
**Preconditions:** Let the Outlook session sit idle until Office prompts re-authentication (may require several hours).
1. Attempt to export after Office session has timed out.

**Expected:** Add-in handles the Office session expiry gracefully — either re-authenticates silently, or prompts the user clearly. Does not crash or show a raw error object.

---

### TC-STAB-06: Add-in on slow network (3G throttle)
**Preconditions:** Use DevTools to throttle network to "Slow 3G" (~400 Kbps).
1. Load task pane.
2. Attempt to export a 5-message thread.

**Expected:** Task pane loads (may exceed 3s target — note actual time). Export completes. No timeouts or crashes.

---

### TC-STAB-07: Large email body
**Preconditions:** An email with a very large HTML body (e.g. newsletter, ≥ 500 KB HTML).
1. Export a thread containing this email.

**Expected:** Export completes. OneNote page is created. HTML sanitiser does not crash on large input. Performance may be slower — note actual time.

---

### TC-STAB-08: Email with no body
**Preconditions:** An email with an empty body (body is `null` or empty string).
1. Export a thread containing this email.

**Expected:** Export completes. The page for that email has the header table (From/To/Date/Subject) and an empty body section — no crash or missing page.

---

### TC-STAB-09: Email with no recipients (CC empty)
**Preconditions:** An email with no CC recipients.
1. Export that thread.

**Expected:** Page header shows CC field as empty or omitted. No crash.

---

## Section 9 — Manifest & Distribution Readiness

### TC-MANIFEST-01: Office Add-in Validator passes
**Preconditions:** Production build complete.
1. Run `npx office-addin-manifest validate manifest.json`.

**Expected:** No errors. Warnings reviewed and addressed or documented.

---

### TC-MANIFEST-02: Single ribbon button only
**Preconditions:** Add-in sideloaded.
1. Examine Home tab and Message tab ribbons.

**Expected:** Exactly one add-in button: "Export to OneNote". No additional buttons.

---

### TC-MANIFEST-03: Task pane opens from ribbon button
**Preconditions:** Add-in sideloaded, email selected.
1. Click "Export to OneNote" ribbon button.

**Expected:** Task pane opens. Correct URL loaded.

---

## Appendix A — Test Data Requirements

| Data needed | Notes |
|-------------|-------|
| Thread with 3–5 messages | Basic export tests |
| Thread with 10 messages | Reliability tests |
| Thread with 20 messages | Performance TC-PERF-02 |
| Thread with 50 messages | Limit TC-LIMIT-01, TC-PERF-03 |
| Thread with 51+ messages | Limit TC-LIMIT-02 |
| Thread with mixed Inbox + Sent messages | TC-EXP-12 |
| Email with attachments | TC-EXP-04 |
| Email with rich HTML body (bold, lists) | TC-EXP-03 |
| Email with large HTML body (≥500 KB) | TC-STAB-07 |
| Email with empty body | TC-STAB-08 |
| Notebook that has been deleted | TC-NB-04 |

---

## Appendix B — Pass/Fail Log Template

```
Date:
Tester:
Environment (OWA-MSA / OWA-Entra / NEW-DT):
Build version:

| TC ID         | Result | Notes |
|---------------|--------|-------|
| TC-AUTH-01    |        |       |
| TC-AUTH-02    |        |       |
| ...           |        |       |
```

---

*Last updated: 2026-02-28*
*Automated E2E (Playwright) planned for future release — see TASKS.md T-F01*
