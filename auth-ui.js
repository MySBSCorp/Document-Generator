// ---------------------------------------------------------------------------
// DOM wiring for the auth gate (#authGate) and nav user badge (#authUser).
// Owns UI state only — all MSAL/token logic lives in auth.js. Never reads
// or displays an access/ID token itself: only account.name and the account's
// email/UPN (via auth.js#getAccountEmail) are shown, and only via
// textContent so a crafted name/claim value can't inject markup into the page.
//
// login()/logout() use MSAL's redirect flow (see auth.js), so a click on
// #loginBtn/#logoutBtn navigates this whole page away — there is no
// "then show the authenticated UI" step to run here for those. The account
// (or an AccessDeniedError) shows up in restoreSession() on the next
// page load instead, once the browser lands back here.
// ---------------------------------------------------------------------------
import { restoreSession, login, logout, getAccountEmail, AccessDeniedError } from "./auth.js";
import { msalConfigured } from "./authConfig.js";

const authGate = document.getElementById("authGate");
const loginBtn = document.getElementById("loginBtn");
const authRetryBtn = document.getElementById("authRetryBtn");
const authErrorMessage = document.getElementById("authErrorMessage");
const authUser = document.getElementById("authUser");
const authAvatar = document.getElementById("authAvatar");
const authUserName = document.getElementById("authUserName");
const authUserEmail = document.getElementById("authUserEmail");
const logoutBtn = document.getElementById("logoutBtn");

const authBlockedModal = document.getElementById("authBlockedModal");
const authBlockedMessage = document.getElementById("authBlockedMessage");
const authBlockedRetryBtn = document.getElementById("authBlockedRetryBtn");

// #authGate visually covers these with a higher z-index, but z-index/paint
// order has no effect on focusability or on whether a bound click/keydown
// handler fires — only `inert` (or `hidden`/`disabled`/removal from the
// DOM) actually stops keyboard/programmatic activation. Security review
// confirmed that without this, Tab could reach the nav's History/Generator
// links behind the overlay and open locally-stored document history with
// no sign-in at all. These start `inert` in the HTML; toggled here so
// authenticated users get the normal, fully-interactive app.
const introOverlay = document.getElementById("introOverlay");
const topnav = document.querySelector(".topnav");
const mainStage = document.getElementById("mainStage");
const historyView = document.getElementById("historyView");
const inertElements = [introOverlay, topnav, mainStage, historyView].filter(Boolean);

function setAppInert(isInert) {
  for (const el of inertElements) {
    el.inert = isInert;
  }
}

function setGateState(state, message) {
  authGate.dataset.state = state;
  if (state === "error") {
    authErrorMessage.textContent = message || "Please try again.";
  }
}

// Pop-up warning shown over the sign-in screen for a disallowed account.
// The sign-in screen underneath is untouched — "Retry" just dismisses this
// and leaves it visible, it does not re-trigger sign-in on its own.
function showBlockedModal(message) {
  authBlockedMessage.textContent = message || "This account isn't authorized to use this application.";
  authBlockedModal.hidden = false;
}

function hideBlockedModal() {
  authBlockedModal.hidden = true;
}

// First letter of up to the first two words (e.g. "Jane Doe" -> "JD"),
// falling back to "?" for an empty/unusable name. textContent-only, so
// this can't inject markup even from an adversarially-crafted name.
function initialsFor(name) {
  const parts = (name || "").trim().split(/\s+/).filter(Boolean);
  if (!parts.length) return "?";
  return parts.slice(0, 2).map((part) => part[0].toUpperCase()).join("");
}

function showAuthenticated(account) {
  authGate.classList.add("is-hidden");
  setAppInert(false);
  const displayName = account?.name || "Signed in";
  authUserName.textContent = displayName;
  authUserEmail.textContent = getAccountEmail(account) || "";
  authAvatar.textContent = initialsFor(displayName);
  authUser.hidden = false;
}

function showUnauthenticated() {
  authUser.hidden = true;
  setAppInert(true);
  authGate.classList.remove("is-hidden");
  setGateState("unauthenticated");
}

async function init() {
  if (!msalConfigured) {
    // authConfig.js still has placeholder client/tenant ID values — don't
    // let MSAL attempt (and fail) a real sign-in against those.
    setGateState(
      "error",
      "Authentication isn't configured yet. Set VITE_MSAL_CLIENT_ID and VITE_MSAL_TENANT_ID (see .env.example / AUTH.md)."
    );
    authGate.classList.remove("is-hidden");
    return;
  }
  try {
    const account = await restoreSession();
    if (account) {
      showAuthenticated(account);
    } else {
      showUnauthenticated();
    }
  } catch (error) {
    if (error instanceof AccessDeniedError) {
      // Entra itself refused this sign-in (see auth.js's AccessDeniedError
      // comment — enforced via Entra's "Assignment required" setting, not
      // any list in this codebase). Show the normal sign-in screen
      // underneath, with the warning pop-up over it.
      showUnauthenticated();
      showBlockedModal(error.message);
      return;
    }
    console.error("[auth] session restore failed:", error);
    setGateState("error");
    authGate.classList.remove("is-hidden");
  }
}

// login() navigates the whole page to Entra and back (see auth.js) — on
// success this function never returns to its caller, the browser just
// leaves the page. Only a same-page failure (e.g. misconfiguration) is
// left to handle here.
async function attemptLogin(options) {
  try {
    await login(options);
  } catch (error) {
    console.error("[auth] login failed:", error);
    setGateState("error", "Sign-in failed. Please try again.");
  }
}

loginBtn.addEventListener("click", async () => {
  loginBtn.disabled = true;
  await attemptLogin();
  loginBtn.disabled = false;
});

authBlockedRetryBtn.addEventListener("click", () => {
  // Just dismiss the pop-up back to the ordinary sign-in screen — this
  // does not itself retry sign-in. The user clicks "Sign in with
  // Microsoft" again from there (Entra's own credential prompt, forced by
  // loginRequest.prompt: "login", is where they'd pick a different account).
  hideBlockedModal();
});

authRetryBtn.addEventListener("click", () => {
  setGateState("loading");
  init();
});

logoutBtn.addEventListener("click", async () => {
  logoutBtn.disabled = true;
  try {
    // Navigates away to Entra's logout endpoint and back on success (see
    // auth.js) — that round trip is itself a full page reload, which also
    // clears any in-memory app state (e.g. this app's extracted W-9 data)
    // along with the auth session. The explicit reload below only runs if
    // logout() fails without navigating anywhere.
    await logout();
  } catch (error) {
    console.error("[auth] logout failed:", error);
    window.location.reload();
  }
});

init();
