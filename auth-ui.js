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
// (or an UnauthorizedEmailError) shows up in restoreSession() on the next
// page load instead, once the browser lands back here.
// ---------------------------------------------------------------------------
import { restoreSession, login, logout, getAccountEmail, UnauthorizedEmailError } from "./auth.js";
import { msalConfigured } from "./authConfig.js";

const authGate = document.getElementById("authGate");
const loginBtn = document.getElementById("loginBtn");
const authRetryBtn = document.getElementById("authRetryBtn");
const authErrorMessage = document.getElementById("authErrorMessage");
const authBlockedMessage = document.getElementById("authBlockedMessage");
const authSwitchAccountBtn = document.getElementById("authSwitchAccountBtn");
const authUser = document.getElementById("authUser");
const authAvatar = document.getElementById("authAvatar");
const authUserName = document.getElementById("authUserName");
const authUserEmail = document.getElementById("authUserEmail");
const logoutBtn = document.getElementById("logoutBtn");

function setGateState(state, message) {
  authGate.dataset.state = state;
  if (state === "error") {
    authErrorMessage.textContent = message || "Please try again.";
  }
  if (state === "blocked") {
    authBlockedMessage.textContent = message || "This account isn't authorized to use this application.";
  }
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
  const displayName = account?.name || "Signed in";
  authUserName.textContent = displayName;
  authUserEmail.textContent = getAccountEmail(account) || "";
  authAvatar.textContent = initialsFor(displayName);
  authUser.hidden = false;
}

function showUnauthenticated() {
  authUser.hidden = true;
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
    if (error instanceof UnauthorizedEmailError) {
      // A sign-in redirect just landed back here with an account outside
      // the allowed domain. Distinct from a generic failure.
      authUser.hidden = true;
      authGate.classList.remove("is-hidden");
      setGateState("blocked", `"${error.email}" isn't authorized to use this application.`);
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

authSwitchAccountBtn.addEventListener("click", async () => {
  authSwitchAccountBtn.disabled = true;
  // Force Entra's account chooser instead of silently retrying whatever
  // Microsoft session the browser still has (which would just get blocked
  // again immediately for the same reason).
  await attemptLogin({ prompt: "select_account" });
  authSwitchAccountBtn.disabled = false;
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
