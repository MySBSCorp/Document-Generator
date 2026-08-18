// ---------------------------------------------------------------------------
// MSAL.js authentication controller for this SPA.
//
// Uses Microsoft's own MSAL Browser library for all OAuth 2.0 / OIDC
// protocol handling (Authorization Code Flow + PKCE, token caching, silent
// renewal). Nothing in this file talks to Entra's authorize/token endpoints
// directly, parses redirect URL fragments, or stores raw OAuth responses —
// that is exactly the kind of hand-rolled OAuth logic MSAL exists to
// replace, and doing it manually is a common source of SPA auth bugs.
//
// Only the standard OIDC scopes (openid/profile/email — see authConfig.js)
// are requested, so this app never acquires or holds a Microsoft Graph (or
// any other API) access token; identity is read straight off the ID token's
// claims after sign-in.
//
// Interactive sign-in/sign-out use the REDIRECT flow (loginRedirect /
// logoutRedirect), not popups. Popups here would rely on this page polling
// `popupWindow.closed` to detect completion/cancellation, but
// login.microsoftonline.com sends its own Cross-Origin-Opener-Policy
// header, which blocks that cross-origin check — MSAL then misreads a
// perfectly normal in-progress sign-in as the popup having been closed and
// throws a spurious `user_cancelled`. That's a browser/COOP-level conflict
// with no popup-side fix. Redirect sidesteps it entirely, and costs nothing
// here: interactive sign-in only ever happens before the wizard is
// reachable (auth-ui.js's gate blocks everything pre-auth), so there's no
// in-memory app state a full-page navigation could lose.
//
// Vendored locally (vendor/msal-browser-4.30.0.esm.js, see vendor/README.md
// for source + SHA-512) rather than imported live from a CDN, for the same
// reason docx/pizzip are vendored: dynamic ES module `import` has no
// browser-native Subresource Integrity mechanism, and CSP `script-src
// 'self'` means the browser only ever runs what's committed to this repo.
// ---------------------------------------------------------------------------
import {
  PublicClientApplication,
  InteractionRequiredAuthError,
  EventType,
} from "./vendor/msal-browser-4.30.0.esm.js";
import { msalConfig, loginRequest, isAllowedEmail } from "./authConfig.js";

export const msalInstance = new PublicClientApplication(msalConfig);

// Thrown when Entra sign-in succeeds but the account's email/UPN isn't in
// authConfig.js's allowed domain. Carries no token/secret data — just the
// (already-not-secret) email that was rejected, for display.
export class UnauthorizedEmailError extends Error {
  constructor(email) {
    super(`"${email}" is not authorized to use this application.`);
    this.name = "UnauthorizedEmailError";
    this.email = email;
  }
}

// The account's email/UPN, read straight off its cached ID token claims —
// no Graph (or any other) API call needed, since none is authorized (only
// openid/profile/email are requested). "email" is the OIDC claim minted
// because that scope was requested; username (the account's UPN) is always
// present for AAD work/school accounts and equals the sign-in email for
// most tenants, so it's used as the fallback for tenants that don't emit
// an "email" claim.
function getAccountEmail(account) {
  return account?.idTokenClaims?.email || account?.username || null;
}

function isEmailAllowed(account) {
  return isAllowedEmail(getAccountEmail(account));
}

// Keeps MSAL's "active account" pointer correct as accounts sign in/out,
// so later acquireTokenSilent() calls don't need an explicit account
// argument. Purely bookkeeping — carries no token material.
msalInstance.addEventCallback((event) => {
  if (
    (event.eventType === EventType.LOGIN_SUCCESS ||
      event.eventType === EventType.ACQUIRE_TOKEN_SUCCESS) &&
    event.payload?.account
  ) {
    msalInstance.setActiveAccount(event.payload.account);
  }
});

let initialized = false;
// Set once, on the page load that lands back here after loginRedirect() —
// holds the account from that redirect response so restoreSession() can
// allow-list-check it before trusting it. Consumed (nulled) after first
// read so it's never re-applied to a later, unrelated restoreSession() call.
let pendingRedirectAccount;
async function ensureInitialized() {
  if (initialized) return;
  await msalInstance.initialize();
  const result = await msalInstance.handleRedirectPromise();
  pendingRedirectAccount = result?.account ?? null;
  initialized = true;
}

export function getActiveAccount() {
  return msalInstance.getActiveAccount() ?? msalInstance.getAllAccounts()[0] ?? null;
}

// Restores whatever session Entra ID itself still considers valid:
// - a fresh account just returned via loginRedirect (checked first), or
// - MSAL's cache (sessionStorage, see authConfig.js), backed by a silent
//   iframe check against Entra's own session when the cache is empty
//   (ssoSilent-style behavior baked into acquireTokenSilent)
// before falling back to "show the sign-in screen". No password or token
// is ever read from anywhere this app controls other than MSAL's own cache.
export async function restoreSession() {
  await ensureInitialized();

  if (pendingRedirectAccount) {
    const account = pendingRedirectAccount;
    pendingRedirectAccount = null;
    if (!isEmailAllowed(account)) {
      // Don't round-trip through logoutRedirect here — that would bounce
      // the user through Entra's logout endpoint and back again just to
      // show a message we can already show on this same page load. Local
      // cleanup is enough: every function in this file re-checks the
      // allow-list before doing anything with an account.
      msalInstance.setActiveAccount(null);
      throw new UnauthorizedEmailError(getAccountEmail(account) || "unknown account");
    }
    msalInstance.setActiveAccount(account);
    return account;
  }

  const account = getActiveAccount();
  if (!account) return null;
  if (!isEmailAllowed(account)) {
    // A cached, Entra-valid session exists for an account outside the
    // allowed domain (e.g. a previous browser session). No interactive
    // step is appropriate on a plain page load, so just stop treating this
    // account as signed in.
    msalInstance.setActiveAccount(null);
    return null;
  }
  try {
    // Only the request's `scopes` are used here, not the full loginRequest
    // — silent requests don't take a `prompt` value.
    await msalInstance.acquireTokenSilent({ scopes: loginRequest.scopes, account });
    return account;
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError) {
      // Cached account exists but Entra says interaction is required
      // (e.g. revoked session, expired refresh token, conditional access).
      // Treat as signed out rather than forcing a login prompt on load.
      return null;
    }
    throw error;
  }
}

// Navigates the whole page to Entra's sign-in page and back — see the
// module comment for why redirect is used instead of a popup. Never
// resolves normally: on success the browser leaves this page entirely;
// the eventual account is picked up by restoreSession() on the page load
// that follows the redirect back.
//
// options.prompt overrides loginRequest's default "login" prompt — e.g.
// pass "select_account" to force Entra's account chooser instead, used by
// the UI's "use a different account" action after a blocked sign-in.
export async function login(options = {}) {
  await ensureInitialized();
  await msalInstance.loginRedirect({ ...loginRequest, ...options });
}

// Navigates to Entra's logout endpoint and back to postLogoutRedirectUri
// (this app's own origin, see authConfig.js) — a full page reload either
// way, which also clears any in-memory app state (e.g. this app's
// extracted W-9 data) along with the auth session.
export async function logout() {
  await ensureInitialized();
  const account = getActiveAccount();
  await msalInstance.logoutRedirect({ account });
}

export { getAccountEmail };
