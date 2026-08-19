// ---------------------------------------------------------------------------
// Microsoft Entra ID (Azure AD) configuration for MSAL.js.
//
// EVERYTHING IN THIS FILE IS PUBLIC. It ships to every browser that loads
// this site and can be read by anyone via "View Source" or devtools. That is
// expected and safe for the three values below:
//
//   - clientId: identifies *which application* is asking to sign someone in.
//     It is not a secret — Microsoft's own docs call it a "public client
//     identifier". Entra ID enforces trust through the registered redirect
//     URI + the fact that this app is registered as a "Single-page
//     application" (public client, no client secret / PKCE-only), not
//     through hiding this value.
//   - tenantId / authority: which Entra directory users sign into.
//   - redirectUri: where Entra is allowed to send the user back to after
//     sign-in. Entra only honors redirect URIs that are explicitly
//     registered on the App Registration, so listing it here does not
//     grant anything by itself.
//
// NEVER put any of the following in this file, anywhere else in this repo,
// or in any file that ends up in the browser bundle:
//   - a client secret / certificate ("Certificates & secrets" in Entra)
//   - a database connection string or API key for a *confidential* backend
//   - any value the Microsoft docs describe as needing to stay server-side
//
// If a future requirement needs a client secret (e.g. calling Graph with
// application permissions instead of delegated user permissions), that
// requires a confidential-client backend to hold the secret — it can never
// be added to this frontend-only app safely. See AUTH.md, Section E.
//
// Plain JS, not TypeScript — this app has no build step, so there's no
// compiler to check type annotations (see AUTH.md's npm/Vite variant if you
// want to add TypeScript alongside a bundler).
// ---------------------------------------------------------------------------

// import.meta.env only exists when a bundler (e.g. Vite, see AUTH.md's
// npm/Vite variant and .env.example) has processed this file; in the
// default no-build setup import.meta.env is undefined and every line below
// falls straight through to the hardcoded placeholder/derived value.
// Optional chaining means this is safe either way — no bundler required.
const env = import.meta.env;

// Application (client) ID from the Entra App Registration's Overview page.
// Safe to expose (see comment above). Hardcoded here (rather than left as
// a placeholder) since this app has no build step to read .env at runtime —
// see .env for the same value, kept there only as an untracked reference.
// env?.VITE_MSAL_CLIENT_ID still overrides this if a bundler is adopted later.
const clientId = env?.MSAL_CLIENT_ID || "bc656f7d-734e-42cc-865f-94afb6497fcd";

// Directory (tenant) ID from the same page. A specific tenant ID (not
// "common"/"organizations") so only accounts in this directory can sign
// in — least-privilege for a single-org app.
const tenantId = env?.MSAL_TENANT_ID || "0c94296c-b6cd-4b8c-a60b-972d913ca913";

// >>> REPLACE with the exact URL this site is served from, with no path,
// query string, or trailing content beyond what you registered in Entra.
// This MUST byte-for-byte match a "Single-page application" redirect URI
// registered on the App Registration (see AUTH.md, Section A). Defaults to
// the page's own origin, which is correct for almost every deployment.
//   Local dev example:  "http://localhost:5500"
//   Production example: "https://your-domain.example"
const redirectUri = env?.MSAL_REDIRECT_URI || window.location.origin;

// True once real values are set in .env / VITE_MSAL_* (or hardcoded below).
// Lets auth-ui.js show a clear "not configured" message instead of letting
// MSAL fail against the placeholder GUID below.
export const msalConfigured = Boolean(clientId && tenantId);

export const msalConfig = {
  auth: {
    clientId: clientId || "00000000-0000-0000-0000-000000000000",
    // Single-tenant authority: only accounts in tenantId's directory can
    // sign in. Change to "https://login.microsoftonline.com/organizations"
    // (any work/school account) or "/common" (+ personal accounts) only if
    // you deliberately want that broader audience.
    authority: `https://login.microsoftonline.com/${tenantId || "common"}`,
    redirectUri,
    // Where MSAL sends the user after logout(). Kept same-origin so the
    // post-logout landing page is this same app, not an external URL.
    postLogoutRedirectUri: redirectUri,
    // Authorization Code Flow + PKCE is what MSAL.js uses by default for
    // SPAs since v2 — there is no separate flag to opt into it, and MSAL
    // has no supported code path for the deprecated implicit grant, so
    // there is nothing to disable here.
    navigateToLoginRequestUrl: true,
  },
  cache: {
    // sessionStorage over localStorage: tokens/cache entries are cleared
    // when the browser tab closes rather than persisting indefinitely on
    // disk. Trade-off: signing in on one tab does not silently authenticate
    // other tabs, and closing the tab means the user must sign in again
    // next visit (mitigated by acquireTokenSilent's iframe-based SSO against
    // Entra's own session, see auth.js). Both storage options are equally
    // readable by any JavaScript running on the page (i.e. by an XSS bug),
    // so this is not a defense against XSS — it only limits how long a
    // stolen-cache-entry window stays open and scopes it to one tab.
    cacheLocation: "sessionStorage",
  },
  system: {
    loggerOptions: {
      // MSAL's own logger. Deliberately does NOT log token values — only
      // protocol-level messages (e.g. "acquireTokenSilent succeeded"). Never
      // change piiLoggingEnabled to true in production: MSAL's PII logging
      // can include account identifiers/claims, which then end up in the
      // browser console and any console-capturing telemetry.
      loggerCallback: (level, message, containsPii) => {
        if (containsPii) return;
        if (level === msalLogLevelError) console.error("[MSAL]", message);
      },
      piiLoggingEnabled: false,
    },
  },
};

// Matches MSAL's LogLevel.Error (0) without importing the enum just for
// this comparison, keeping this config file dependency-free.
const msalLogLevelError = 0;

// Minimum viable scopes: these are the standard OIDC scopes needed only to
// sign the user in and read their name/email/UPN off the ID token — no
// Microsoft Graph (or any other API) permission is requested at all, so
// there's no access token to mismanage and nothing beyond identity to
// misuse even if a token were somehow leaked.
//
// prompt: "login" forces Entra to show the credential prompt every time,
// ignoring any existing browser session, so an interactive sign-in never
// silently reuses whatever Microsoft account happens to already be logged
// into the browser. Trade-off: less convenient than true single sign-on
// (restoreSession()'s silent/cached-session path in auth.js is unaffected
// by this — it only applies to the interactive loginRedirect() call).
export const loginRequest = {
  scopes: ["openid", "profile", "email"],
  prompt: "login",
};

// ---------------------------------------------------------------------------
// No allow-list of any kind lives in this file, or anywhere else in this
// frontend. That list used to be a hardcoded array of 5 email addresses
// here — removed deliberately, because anything present in code that ships
// to the browser can always be read via "View Source"/devtools, no matter
// how it's stored (hardcoded, in a fetched JSON file, Base64-encoded,
// whatever). There is no way to make client-side JavaScript keep a secret
// from the browser running it.
//
// Who may sign in is instead enforced entirely by Entra ID itself, outside
// this codebase: Entra admin center → Enterprise applications → this app
// → Properties → "Assignment required?" = Yes, then assign exactly the
// intended users under "Users and groups". Entra refuses sign-in for
// anyone not assigned *before* it ever hands a token back to this app —
// the actual list of who's allowed lives only in the Entra admin portal,
// visible only to whoever has admin access there. See AUTH.md, Section A
// (setup steps) and Section E (why this is the only way to truly keep this
// list out of the frontend).
// ---------------------------------------------------------------------------
