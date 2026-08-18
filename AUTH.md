# Authentication (Microsoft Entra ID + MSAL.js)

Frontend-only sign-in for this static site using [MSAL.js](https://github.com/AzureAD/microsoft-authentication-library-for-js)
(`@azure/msal-browser`) and the OAuth 2.0 Authorization Code Flow with PKCE.
There is no backend, database, or server component, and none is required
for authentication — MSAL.js implements the full browser-side protocol.
Sign-in only requests the standard OIDC scopes (`openid`/`profile`/`email`)
to read the user's identity off their ID token — no Microsoft Graph (or any
other API) permission is requested at all.

## A. Microsoft Entra configuration

Create the App Registration once, in the Entra tenant that should own this app:

1. **Entra admin center** → *Identity* → *Applications* → *App registrations* → **New registration**.
2. **Name**: anything recognizable (e.g. `Document Generator`).
3. **Supported account types**: *Accounts in this organizational directory only (Single tenant)* — matches `authConfig.js`'s single-tenant `authority`. Pick *Multitenant* only if you deliberately want other organizations' users to sign in too.
4. **Redirect URI**: platform **Single-page application (SPA)** — not "Web". This is required: the SPA platform is what tells Entra to use Authorization Code Flow + PKCE for a public client with no client secret, instead of a confidential-client flow. Add the exact origin(s) this site is served from, no path/trailing slash beyond what you'll pass as the redirect URI:
   - Local dev: `http://localhost:5500` (or whatever port your static file server uses)
   - Production: `https://your-production-domain.example`
5. **Register**.
6. On the app's **Overview** page, copy:
   - **Application (client) ID** → `VITE_MSAL_CLIENT_ID` (see Section F)
   - **Directory (tenant) ID** → `VITE_MSAL_TENANT_ID`
7. **Authentication** blade: confirm the SPA platform and redirect URI(s) are listed. Leave "ID tokens" / "Access tokens" (implicit grant) checkboxes **unchecked** — this app uses the auth code + PKCE flow the SPA platform sets up automatically, not the implicit flow.
8. **API permissions**: leave the default (`Microsoft Graph → User.Read` delegated, added automatically by most tenants) or remove it entirely — this app never requests it and calls no Graph endpoint (see [Section E](#e-security-explanation)). Nothing to add here.
9. **Certificates & secrets**: leave empty. A SPA public client must never have a client secret — if one exists here from experimenting, delete it.
10. *(Optional, stronger enforcement)* This app also enforces its own domain allow-list at the application level (`authConfig.js`'s `ALLOWED_EMAIL_DOMAIN`, checked in `auth.js`) — see [Section E](#e-security-explanation). If you'd rather Entra itself refuse sign-in for anyone outside that domain (so a blocked user can't even complete the Microsoft prompt), go to **Enterprise applications** → this app → **Properties** → set **"Assignment required?"** to **Yes**, then assign only the intended users/group. The two mechanisms are independent and can be used together.

Nothing else in Entra needs configuring for this app.

## B. Project folder structure

```
Document-Generator/
├── index.html            # markup: auth gate, nav user badge, app UI
├── style.css             # includes #authGate / .auth-user styles
├── authConfig.js         # PUBLIC config: client ID, tenant ID, redirect URI, scopes, allowed domain
├── auth.js               # MSAL controller: init, login, logout, silent session restore
├── auth-ui.js            # DOM wiring between auth.js and index.html
├── .env.example          # documents VITE_MSAL_CLIENT_ID / VITE_MSAL_TENANT_ID / VITE_MSAL_REDIRECT_URI
├── script.js             # existing app logic (unrelated to auth)
├── vendor/
│   ├── msal-browser-4.30.0.esm.js   # vendored MSAL.js (see vendor/README.md)
│   ├── msal-common-15.17.0.esm.js   # MSAL's dependency, also vendored
│   └── README.md                    # source URLs + SHA-512 for every vendored file
└── AUTH.md               # this file
```

No `node_modules`, build step, or bundler is required to run this. If you'd
rather manage MSAL via npm/Vite (e.g. to get automatic updates through a
lockfile instead of manually re-vendoring), see [Section G](#npmvite-alternative).

## C. Complete code

All the auth-specific files above are already in the repo:

- [`authConfig.js`](authConfig.js) — the only file you edit per environment (Section F). Exports `msalConfig`, `loginRequest`, `msalConfigured`, `ALLOWED_EMAIL_DOMAIN`, and `isAllowedEmail(email)`.
- [`auth.js`](auth.js) — wraps `PublicClientApplication`: `restoreSession()`, `login()`, `logout()`, `getAccountEmail(account)`. No Graph/API token acquisition exists in this file at all.
- [`auth-ui.js`](auth-ui.js) — shows/hides `#authGate`'s loading/unauthenticated/blocked/error panels and the nav's `#authUser` badge; only ever displays `account.name` and `getAccountEmail(account)`, always via `textContent`.
- [`index.html`](index.html) — `#authGate` overlay (z-index above the intro animation, so nothing is reachable pre-auth) and the `#authUser` nav badge; updated CSP.
- [`style.css`](style.css) — `#authGate` / `.auth-user` styles (flat, full-screen layout — no boxed card).
- [`.env.example`](.env.example) — documents the three `VITE_MSAL_*` values; copy to `.env` if using the [npm/Vite variant](#npmvite-alternative).
- [`vendor/msal-browser-4.30.0.esm.js`](vendor/msal-browser-4.30.0.esm.js) + [`vendor/msal-common-15.17.0.esm.js`](vendor/msal-common-15.17.0.esm.js) — MSAL itself, vendored (not loaded from a live CDN) for the same Subresource-Integrity reasoning already used for this project's other ES-module dependencies; see `vendor/README.md`.

## D. How authentication works

1. **Page load** (`auth-ui.js` → `auth.js#restoreSession`): if `authConfig.js`'s `msalConfigured` is `false` (placeholder client/tenant ID still in place), the gate shows a "not configured" message and stops there. Otherwise `PublicClientApplication.initialize()` runs, then `handleRedirectPromise()` — a no-op on a plain page load, but on the page load that lands back here right after a sign-in redirect (step 4), this is what resolves with the new account. If no fresh redirect result is pending, MSAL's cache (`sessionStorage`) is checked for an existing account instead. If one exists and its email/UPN is on the allowed domain, `acquireTokenSilent()` is called — MSAL either serves a cached token, silently refreshes it, or performs a hidden-iframe check against Entra's own session, with no user-visible interaction.
2. **Silent success** → the account is treated as signed in, the auth gate hides, and the nav badge shows `account.name` / `getAccountEmail(account)` — read directly from the cached ID token, no network call.
3. **Silent failure** (`InteractionRequiredAuthError`, e.g. no cached session, revoked consent, expired refresh token) → the gate shows the "Sign in with Microsoft" button. Nothing else in the app is reachable at this point (Section E).
4. **Login** (`auth.js#login`): clicking the button calls `loginRedirect()` with `prompt: "login"` (see `authConfig.js`'s `loginRequest`), which forces Entra's credential prompt every time. This navigates the *entire page* to Entra's `/authorize` endpoint — there is no popup — runs Authorization Code Flow + PKCE, and Entra redirects the browser back to this app's `redirectUri` once sign-in completes. Redirect (not popup) is used deliberately: see the module comment at the top of `auth.js` for why popups and Microsoft's own Cross-Origin-Opener-Policy header don't reliably work together, and why redirect costs nothing here (sign-in only ever happens before the wizard is reachable).
5. **Domain check** (`auth.js#isEmailAllowed`), on the page load that follows step 4: the returned account's email/UPN (read from `idTokenClaims.email`, falling back to `username`) is checked against `authConfig.js`'s `ALLOWED_EMAIL_DOMAIN` via `isAllowedEmail()`. If it doesn't match, `restoreSession()` clears the local active-account pointer and throws `UnauthorizedEmailError`; `auth-ui.js` catches that specifically (right there in `init()`, same page load) and shows the "Access restricted" panel with a "use a different account" option (which re-opens sign-in with `prompt: "select_account"`), distinct from a generic sign-in failure. Only on a match does the account become MSAL's "active account" and step 2's authenticated flow run.
6. **Logout** (`auth.js#logout`): `logoutRedirect()` navigates to Entra's logout endpoint and back to `postLogoutRedirectUri` (this app's own origin) — that round trip is itself a full page reload, which also clears any in-memory app state (e.g. this app's extracted W-9 data) along with the auth session.

## E. Security explanation

- **No secrets in the frontend, ever.** `authConfig.js` contains a client ID, tenant ID, and redirect URI — all public by design (Entra's own docs call the client ID a "public client identifier"). There is no client secret anywhere in this codebase, and there cannot be: a SPA is a *public* OAuth client. If a future feature genuinely needs a confidential secret (e.g. calling an API with *application* permissions instead of the user's own delegated permissions), that requires a separate backend to hold it — it can never be safely added to this frontend. Anything shipped in browser JavaScript, `.env` files bundled into a build, Base64, or "obfuscated" code is inspectable by any user; none of that is a substitute for a real backend when a secret is actually required.
- **Auth code + PKCE, not implicit flow.** Registering the app as **SPA** (not "Web") is what makes MSAL.js use Authorization Code Flow with PKCE — MSAL.js has no supported code path for the deprecated implicit grant, so there's nothing to accidentally opt into.
- **No API permissions requested at all.** `loginRequest.scopes` is `["openid", "profile", "email"]` — standard OIDC scopes that only let this app read the signed-in user's own name/email/UPN off their ID token. No Microsoft Graph (or any other API) access token is ever acquired, so there's no broader permission to over-request and no API access token to mismanage, log, or leak.
- **Redirect, not popup, for interactive sign-in.** `login()`/`logout()` use `loginRedirect()`/`logoutRedirect()`. Popup-based flows need this page to poll `popupWindow.closed` to detect success/cancellation, but `login.microsoftonline.com` sends its own Cross-Origin-Opener-Policy header, which blocks that cross-origin check in current browsers — MSAL then misreads an in-progress sign-in as a closed/cancelled popup (`user_cancelled`), even though nothing was cancelled. Redirect has no such dependency. The trade-off (a full-page navigation instead of a popup) costs nothing in this app specifically, since interactive sign-in only ever happens before the wizard is reachable — there's no in-memory app state yet for a navigation to lose.
- **Forced re-authentication, by design.** `loginRequest.prompt: "login"` makes every interactive sign-in show Entra's credential prompt rather than silently reusing whatever Microsoft session the browser already has — useful on a shared/kiosk-style machine where "whichever Microsoft account happens to be logged into this browser" should never be auto-accepted. This only affects the interactive `loginRedirect()` call; `restoreSession()`'s silent/cached-session check on page load is unaffected, so a user who already completed this app's own sign-in still gets a smooth reload experience.
- **Cache location trade-off** (`authConfig.js`'s `cache.cacheLocation: "sessionStorage"`): tokens/account data live only as long as the tab is open, rather than persisting on disk indefinitely (`localStorage`). Trade-off: signing in on one tab does not authenticate other tabs, and closing the tab means signing in again next time (mitigated by the silent SSO check in step 1 above). **Neither `sessionStorage` nor `localStorage` is safe from an XSS bug** — any script running on the page can read either. This setting only bounds how long a compromised-page window stays exploitable, and is not a substitute for the CSP/sanitization measures below.
- **Tokens never touch the console, DOM, or URL.** `authConfig.js`'s MSAL logger callback explicitly drops any message flagged as containing PII and never logs at the point tokens are received; `auth-ui.js` never puts a token value into `textContent`/`innerHTML`/an attribute/a URL — it only ever reads `account.name` and ID-token *claims* (name/email), never the raw token strings.
- **Protected-content gate.** `#authGate` is a fixed, full-viewport overlay with a z-index above every other element in the page (including the intro animation), and starts in the `loading` state. The app underneath is not merely visually hidden — with the overlay covering it and blocking interaction until the gate is hidden, nothing in the wizard is usable pre-auth.
- **Output sanitization.** The displayed name/email always go through `textContent`, never `innerHTML`, so even a maliciously-crafted Entra display name or claim value can't inject markup or script into the page — consistent with this app's existing rule (see the CSP comment in `index.html`) of never using `innerHTML` for any value that ultimately comes from outside this codebase.
- **CSP.** `script-src 'self'` (plus the existing SRI-pinned `cdnjs.cloudflare.com` script tags and `'wasm-unsafe-eval'` for Tesseract) means the browser will only ever execute JavaScript committed to this repo — MSAL is vendored locally rather than imported from a live CDN for exactly this reason. `connect-src` is extended only for the one host MSAL actually needs (`login.microsoftonline.com`) — there's no Graph host to add, since this app never calls Graph; `frame-src` is extended only for MSAL's hidden-iframe silent-SSO check against the same host. No `unsafe-inline` or `unsafe-eval` is added anywhere for this feature.
- **No `eval`, no dynamic script execution, no unsafe HTML injection** anywhere in `auth.js`/`auth-ui.js`/`authConfig.js`.
- **Application-level domain allow-list, with an Entra-level option.** `authConfig.js`'s `ALLOWED_EMAIL_DOMAIN` restricts usage to accounts on that domain, enforced client-side in `auth.js` immediately after every sign-in (and on every cached-session restore). This is an *app-level* control — Entra itself still lets any account in the tenant complete sign-in, and this app then signs a non-matching one back out. Anyone who can edit and re-host this frontend's files can also edit that value, so if the actual security boundary needs to survive a compromised/rogue deployment of this code, use the Entra-level "Assignment required" restriction from Section A alongside it (or instead of it) — that one Entra enforces regardless of what the frontend code says.
- **Frontend code is never fully confidential.** Anyone can open devtools and read `authConfig.js`, `auth.js`, and the vendored MSAL source in full. That is expected and accounted for: the only things that keep this app secure are (a) no secret exists in it to steal, and (b) Entra enforces trust via the registered SPA redirect URI + PKCE, not via hiding the client ID.

## F. Local testing instructions

1. Register the app per [Section A](#a-microsoft-entra-configuration), using your local dev URL (e.g. `http://localhost:5500`) as the SPA redirect URI. `http://localhost` origins are allowed by Entra for SPA redirect URIs even without HTTPS; any other non-localhost origin must be HTTPS.
2. Set real values for the three `>>> REPLACE` markers in `authConfig.js`:
   - `clientId` → the Application (client) ID from Entra's Overview page.
   - `tenantId` → the Directory (tenant) ID from the same page.
   - `redirectUri` is already computed from `window.location.origin` — you don't need to hardcode it, but it **must exactly match** a redirect URI registered in Entra (same scheme, host, and port).
   - `ALLOWED_EMAIL_DOMAIN` → the email domain that should be allowed to sign in (defaults to `mysbscorp.com`).

   Either edit `authConfig.js` directly, or (if using the [npm/Vite variant](#npmvite-alternative)) copy [`.env.example`](.env.example) to `.env` and fill in `VITE_MSAL_CLIENT_ID` / `VITE_MSAL_TENANT_ID` / `VITE_MSAL_REDIRECT_URI` there instead.
3. Serve the folder over HTTP — don't open `index.html` via `file://`, since MSAL's redirect flow and the CSP both assume a real origin. Any static server works, e.g.:
   ```bash
   npx serve .
   # or
   python -m http.server 5500
   ```
4. Open the dev URL, click **Sign in with Microsoft**, sign in with an allowed-domain account (the whole page navigates to Entra and back), and confirm the nav shows your name/email and the wizard becomes usable. Reload the page and confirm it signs you back in silently (no prompt) via `restoreSession()`. Click **Sign out** and confirm the app navigates through Entra's logout and back to the sign-in screen. Then try signing in with an account **outside** the allowed domain and confirm you land on the "Access restricted" panel, not the app.

## G. Production deployment instructions

This is a static site — deploy it exactly like the rest of this app (any static host: e.g. Azure Static Web Apps, Netlify, GitHub Pages, an S3 bucket + CDN, or a plain web server serving these files). There is nothing auth-specific about the deployment itself, with two things to get right:

1. **Register the production origin** as an additional SPA redirect URI on the same (or a separate, environment-specific) App Registration, e.g. `https://your-production-domain.example`. Many teams use *separate* App Registrations for dev/staging/prod so a compromised dev redirect URI can't be abused against production — if you do that, the client/tenant ID in `authConfig.js` become per-environment values you swap at deploy time (see below).
2. **Serve over HTTPS.** Entra requires HTTPS for any non-localhost SPA redirect URI, and MSAL's PKCE/token handling assumes a secure context.

### Per-environment configuration without a backend

Since there's no build step, the client/tenant ID are plain values in a committed file — safe, because they're public (Section E). If you want different values per environment without hand-editing the file on each deploy:

- Simplest: keep separate branches/copies of `authConfig.js` per environment and let your deploy pipeline pick the right one, or template it as a build step that only substitutes these public values (still committing no secret).
- If you adopt the [npm/Vite variant](#npmvite-alternative), use Vite's `import.meta.env.VITE_*` build-time variables — but remember these still end up as plain text in the shipped bundle; `VITE_*` variables are not a secrets mechanism, only a build-time templating one.

### NPM/Vite alternative

If a build step is acceptable, replace the vendored `vendor/msal-browser-*.esm.js` with the real npm package instead of manually re-vendoring on updates:

```bash
npm install @azure/msal-browser
```

```js
// authConfig.js — same shape as the CDN-free version above
import { PublicClientApplication } from "@azure/msal-browser";
```

`authConfig.js` already reads `import.meta.env.VITE_MSAL_CLIENT_ID` /
`VITE_MSAL_TENANT_ID` / `VITE_MSAL_REDIRECT_URI` first, falling back to
the hardcoded placeholders when no bundler is present — so adopting Vite
just means copying [`.env.example`](.env.example) to `.env` and filling in
real values; no code changes needed. `.env` is still not a secrets
mechanism here (Section E) — it's build-time templating for public values,
and Vite inlines them into the shipped bundle same as if you'd hardcoded
them.

Everything else (`auth.js`'s API surface, the CSP requirements, the security properties in Section E) is unchanged — Vite just handles bundling/dev-serving instead of a plain static file server. This is optional; the core deliverable in this repo has no build dependency.

## H. Security checklist before deployment

- [ ] `authConfig.js` has real `clientId`/`tenantId` values (or `.env` has real `VITE_MSAL_CLIENT_ID`/`VITE_MSAL_TENANT_ID`), not the placeholder GUID — `msalConfigured` should evaluate to `true`.
- [ ] `authConfig.js`'s `ALLOWED_EMAIL_DOMAIN` is the real, intended domain, not left as a placeholder — otherwise every (or no) sign-in will be blocked/allowed unexpectedly.
- [ ] The App Registration's platform is **Single-page application**, not "Web" or "Mobile and desktop applications".
- [ ] Every origin this app is actually served from (including production) is registered as a SPA redirect URI, exact scheme/host/port match.
- [ ] No client secret exists on the App Registration's "Certificates & secrets" blade.
- [ ] "Implicit grant and hybrid flows" (ID tokens / access tokens checkboxes) are left unchecked on the Authentication blade.
- [ ] No API permission beyond the tenant's default is added — this app calls no API.
- [ ] `grep -R "client_secret\|clientSecret\|CLIENT_SECRET"` (and similarly for API keys/passwords) across the repo returns nothing in frontend-shipped files.
- [ ] The CSP in `index.html` still has no `unsafe-inline`/`unsafe-eval`, and `connect-src`/`frame-src` list only the hosts actually needed.
- [ ] Site is served over HTTPS in production.
- [ ] `vendor/msal-browser-*.esm.js` / `vendor/msal-common-*.esm.js` match the hashes recorded in `vendor/README.md` (re-verify with `sha512sum` if you ever re-fetch them).
- [ ] Manual smoke test: sign in with an allowed-domain account, reload (silent restore works), sign out, sign in with a non-allowed-domain account and confirm it's blocked, and confirm the wizard is not reachable/usable before signing in at all.

## I. Common mistakes to avoid

- **Registering the app as "Web" instead of "SPA".** This silently changes which OAuth flow Entra expects and can produce confusing `AADSTS9002326` ("cross-origin token redemption") errors from a public client with no secret.
- **Putting a client secret in the frontend "just to be safe" or "for extra permissions."** A SPA can't keep a secret confidential no matter how it's stored client-side (env file, obfuscation, encoding) — if a feature needs one, it needs a backend, full stop.
- **Adding Microsoft Graph (or any other API) scopes "in case they're needed later."** Every added scope is something a compromised page (or a user re-consenting under duress) could misuse; this app currently needs none beyond `openid`/`profile`/`email`.
- **Enabling `piiLoggingEnabled` in MSAL's logger**, "just for debugging" — it can leak account claims into browser consoles or any console-capturing telemetry tool.
- **Mismatched redirect URI** (trailing slash, `http` vs `https`, wrong port) — Entra matches SPA redirect URIs exactly; a mismatch fails sign-in with `AADSTS50011`.
- **Testing over `file://`** instead of a real HTTP origin — MSAL's redirect flow and this app's CSP both assume `http(s)://`.
- **Assuming `sessionStorage` cache = safe from XSS.** It only limits *how long* a token lives, not whether an XSS bug can read it — the real defense is the CSP and never using `innerHTML` with untrusted data (see Section E).
- **Treating the domain allow-list as the only line of defense.** It's an app-level convenience check that anyone editing this frontend's code can also edit; use Entra's "Assignment required" (Section A, step 10) if the restriction genuinely needs to survive a compromised deployment.
- **Switching sign-in back to a popup "for a nicer UX"** without checking COOP first. If you ever do move some *other* interactive call back to a popup (e.g. a future `acquireTokenPopup` for some added API scope), be aware `login.microsoftonline.com`'s own COOP header can cause the exact spurious `user_cancelled` this app avoided by using redirect (see Section E) — this isn't fixable from this app's side.
