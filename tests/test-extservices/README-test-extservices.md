Various tests to validate functionality of external services and APIs.


backend-example.js
graph-auth-test.js
graph-get-users.js
interactive-graph-test.js
pkce-debug-interactive.html
pkce-debug-server.js
README-test-extservices.md
simple-graph-test.js
test-pkce-auth.js

## Test descriptions

### `backend-example.js`
- **Type:** Backend sample (Node/Express), not a direct test runner.
- **Purpose:** Demonstrates server-side validation of an Office SSO token and OBO token exchange using `@azure/msal-node`.
- **What it verifies:**
	- Request handling for `POST /api/auth/exchange-token`
	- JWT signature validation via Azure AD JWKS endpoint
	- OBO flow to obtain Graph token
	- Graph call to `/me/onenote/notebooks`
- **When useful:** Prototyping a secure backend-assisted auth flow for add-in scenarios.

### `graph-auth-test.js`
- **Type:** CLI auth smoke test.
- **Purpose:** Validates that the app registration can issue a delegated token via Device Code flow.
- **What it verifies:**
	- `PublicClientApplication` initialization
	- Device code prompt flow
	- Receipt of a usable access token for `User.Read`
- **Output:** Prints a success message and raw token on successful sign-in.

### `graph-get-users.js`
- **Type:** CLI Graph permissions/tenant diagnostic.
- **Purpose:** Tests client-credentials access with app permissions.
- **What it verifies:**
	- App-only token acquisition using `ClientSecretCredential`
	- Access to `/applications` (sanity check endpoint)
	- Access to `/users` and common permission failures (`Authorization_RequestDenied`, `Forbidden`)
- **Output:** Detailed troubleshooting guidance for missing admin consent or wrong permission type.

### `interactive-graph-test.js`
- **Type:** Interactive browser-based integration test.
- **Purpose:** Simulates add-in-like interactive auth in a local Node script.
- **What it verifies:**
	- Browser launch to Azure authorization endpoint
	- Local callback handling on `http://localhost:8080/callback`
	- Access token capture from fragment
	- Graph calls to `/me` and `/me/onenote/notebooks`
- **Output:** User identity and discovered notebook metadata.

### `pkce-debug-interactive.html`
- **Type:** Manual PKCE flow debugger UI.
- **Purpose:** Step-by-step validation tool for PKCE implementation details.
- **What it verifies (by phases):**
	- Crypto helpers and environment config
	- Authorization URL construction (`code_challenge`, `state`, scopes)
	- Callback parsing / code extraction
	- Token exchange and backend exchange fallback
	- Token usage against Graph APIs (`/me`, notebooks)
- **Output:** Rich in-page logs and phase status to isolate PKCE/auth issues quickly.

### `pkce-debug-server.js`
- **Type:** Local helper server for PKCE debugger.
- **Purpose:** Serves static assets and renders a friendly callback page for auth result inspection.
- **What it verifies/supports:**
	- Debug page hosting (`/pkce-debug`)
	- Callback inspection route (`/src/auth/callback.html`)
	- URL/code/state copy-back workflow into debugger
	- Basic health endpoint (`/health`)
- **Note:** Uses `tests/test-extservices/pkce-debug-interactive.html` in the current repo layout.

### `simple-graph-test.js`
- **Type:** CLI Graph endpoint smoke test with manual token.
- **Purpose:** Fast validation of OneNote Graph endpoints using a token from Graph Explorer.
- **What it verifies:**
	- Graph connectivity via `/me`
	- Notebook listing via `/me/onenote/notebooks`
	- Section listing via `/me/onenote/notebooks/{id}/sections`
- **Output:** Notebook/section metadata and guidance for token acquisition.

### `test-pkce-auth.js`
- **Type:** PKCE/auth module functional test suite.
- **Purpose:** Programmatic checks of PKCE helper modules and auth orchestration.
- **What it verifies:**
	- Config validation behavior
	- Crypto generation (`code_verifier`, `code_challenge`, `state`)
	- PKCE authenticator initialization and URL generation
	- Token storage/expiry logic
	- Mock notebook data shape and integration entrypoint behavior
- **Output:** Pass/fail counters with per-test diagnostics.

## Quick usage notes
- These scripts are primarily diagnostics and prototypes for external auth/Graph behavior.
- Scripts in this folder load environment values from the root `.env` via `load-root-env.cjs`.
- Prefer these for troubleshooting authentication and Microsoft Graph permission issues before integrating changes into `src/`.

## Environment variables (root `.env`)

### Used by multiple tests
- `GRAPH_CLIENT_ID` (or existing `VITE_CLIENT_ID` fallback)
- `GRAPH_TENANT_ID` (or existing `VITE_TENANT_ID` fallback)
- `GRAPH_AUTHORITY` (optional; defaults to `https://login.microsoftonline.com/{tenant}`)

### `graph-auth-test.js`
- `GRAPH_DEVICE_CODE_SCOPES` (optional, comma-separated; default: `User.Read`)

### `graph-get-users.js`
- `AZURE_CLIENT_SECRET` (required)
- `GRAPH_APP_SCOPES` (optional, comma-separated; default: `https://graph.microsoft.com/.default`)

### `interactive-graph-test.js`
- `GRAPH_INTERACTIVE_PORT` (optional; default: `8080`)
- `GRAPH_REDIRECT_URI` (optional; default: `http://localhost:{GRAPH_INTERACTIVE_PORT}/callback`)
- `GRAPH_INTERACTIVE_SCOPES` (optional, comma-separated)

### `simple-graph-test.js`
- `GRAPH_ACCESS_TOKEN` (required)

### `backend-example.js`
- `AZURE_CLIENT_ID` (required; falls back from `GRAPH_CLIENT_ID`)
- `AZURE_CLIENT_SECRET` (required)
- `EXTSERVICES_CORS_ORIGIN` (optional; default: `https://localhost:3000`)
- `GRAPH_OBO_SCOPES` (optional, comma-separated)

### `pkce-debug-server.js`
- `PKCE_DEBUG_SERVER_PORT` (optional; fallback `GRAPH_INTERACTIVE_PORT`; default `3000`)
