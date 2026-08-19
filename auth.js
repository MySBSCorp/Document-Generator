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
  ServerError,
  EventType,
} from "./vendor/msal-browser-4.30.0.esm.js";
import { msalConfig, loginRequest } from "./authConfig.js";

export const msalInstance = new PublicClientApplication(msalConfig);

// There is no allowed-account list anywhere in this file, or anywhere else
// in this frontend: who may sign in is enforced entirely by Entra itself
// (Enterprise Applications → this app → "Assignment required" = Yes, with
// exactly the intended users assigned — see AUTH.md, Section A/E). Entra
// refuses the sign-in for anyone not assigned before it ever hands back a
// token, so this app never receives — and never has to hardcode, ship, or
// otherwise expose — the list of who's allowed. In practice, Entra usually
// shows its own hosted error page for this (the same kind of page you'd
// see for a redirect URI mismatch) rather than redirecting back here at
// all; this error class only covers the cases where it does redirect back
// with an error attached, so this app can still show something sensible.
export class AccessDeniedError extends Error {
  constructor(message) {
    super(message || "This account is not authorized to use this application.");
    this.name = "AccessDeniedError";
  }
}

// The account's email/UPN, read straight off its cached ID token claims —
// no Graph (or any other) API call needed, since none is authorized (only
// openid/profile/email are requested). "email" is the OIDC claim minted
// because that scope was requested; username (the account's UPN) is always
// present for AAD work/school accounts and equals the sign-in email for
// most tenants, so it's used as the fallback for tenants that don't emit
// an "email" claim. Purely for display after a successful, Entra-approved
// sign-in — never used here to decide who's allowed in.
function getAccountEmail(account) {
  return account?.idTokenClaims?.email || account?.username || null;
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
// pick it up. Consumed (nulled) after first read so it's never re-applied
// to a later, unrelated restoreSession() call.
let pendingRedirectAccount;
async function ensureInitialized() {
  if (initialized) return;
  await msalInstance.initialize();
  try {
    const result = await msalInstance.handleRedirectPromise();
    pendingRedirectAccount = result?.account ?? null;
  } catch (error) {
    initialized = true;
    // Entra returned an error instead of a token on the redirect back —
    // e.g. error=access_denied when "Assignment required" rejected this
    // account. See AccessDeniedError's comment above.
    if (
      error instanceof ServerError &&
      (error.errorCode === "access_denied" || /AADSTS50105/.test(error.errorMessage || ""))
    ) {
      throw new AccessDeniedError();
    }
    throw error;
  }
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
// is ever read from anywhere this app controls other than MSAL's own
// cache. No allow-list check happens here either — see AccessDeniedError's
// comment above: Entra itself is what decides who gets this far at all.
export async function restoreSession() {
  await ensureInitialized(); // may throw AccessDeniedError — let it propagate

  if (pendingRedirectAccount) {
    const account = pendingRedirectAccount;
    pendingRedirectAccount = null;
    msalInstance.setActiveAccount(account);
    return account;
  }

  const account = getActiveAccount();
  if (!account) return null;
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
// options.prompt overrides loginRequest's default "login" prompt if a
// caller ever needs a different one — e.g. "select_account" to force
// Entra's account chooser. Not currently used by auth-ui.js (its "Retry"
// on a blocked sign-in just dismisses the pop-up; it doesn't call this).
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
