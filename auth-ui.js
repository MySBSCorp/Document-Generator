import { restoreSession, login, logout, getAccountEmail, AccessDeniedError } from "./auth.js";
import { msalConfigured } from "./authConfig.js";

const e2eAuthBypassed = import.meta.env.DEV && import.meta.env.VITE_E2E_BYPASS_AUTH === "true";

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

const authLogoutConfirmModal = document.getElementById("authLogoutConfirmModal");
const authLogoutCancelBtn = document.getElementById("authLogoutCancelBtn");
const authLogoutConfirmBtn = document.getElementById("authLogoutConfirmBtn");
const authLogoutUserInfo = document.getElementById("authLogoutUserInfo");
const authLogoutUserAvatar = document.getElementById("authLogoutUserAvatar");
const authLogoutUserName = document.getElementById("authLogoutUserName");
const authLogoutUserEmail = document.getElementById("authLogoutUserEmail");

const authBackNavConfirmModal = document.getElementById("authBackNavConfirmModal");
const authBackNavCancelBtn = document.getElementById("authBackNavCancelBtn");
const authBackNavConfirmBtn = document.getElementById("authBackNavConfirmBtn");

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

function showBlockedModal(message) {
  authBlockedMessage.textContent = message || "This account isn't authorized to use this application.";
  authBlockedModal.hidden = false;
}

function hideBlockedModal() {
  authBlockedModal.hidden = true;
}

function initialsFor(name) {
  const parts = (name || "").trim().split(/\s+/).filter(Boolean);
  if (!parts.length) return "?";
  return parts.slice(0, 2).map((part) => part[0].toUpperCase()).join("");
}

let isAuthenticated = false;
let backGuardArmed = false;
let currentAccount = null;

function armBackGuard() {
  if (backGuardArmed) return;
  backGuardArmed = true;
  history.pushState({ authGuard: true }, "");
}

function disarmBackGuard() {
  backGuardArmed = false;
}

window.addEventListener("popstate", () => {
  if (!isAuthenticated) return;
  history.pushState({ authGuard: true }, "");
  authBackNavConfirmModal.hidden = false;
});

authBackNavCancelBtn.addEventListener("click", () => {
  authBackNavConfirmModal.hidden = true;
});

authBackNavConfirmBtn.addEventListener("click", async () => {
  authBackNavConfirmBtn.disabled = true;
  authBackNavCancelBtn.disabled = true;
  try {
    await logout();
  } catch (error) {
    console.error("[auth] logout failed:", error);
    window.location.reload();
  } finally {
    authBackNavConfirmBtn.disabled = false;
    authBackNavCancelBtn.disabled = false;
    authBackNavConfirmModal.hidden = true;
  }
});

function showAuthenticated(account) {
  authGate.classList.add("is-hidden");
  setAppInert(false);
  currentAccount = account;
  const displayName = account?.name || "Signed in";
  authUserName.textContent = displayName;
  authUserEmail.textContent = getAccountEmail(account) || "";
  authAvatar.textContent = initialsFor(displayName);
  authUser.hidden = false;
  isAuthenticated = true;
  armBackGuard();
}

function showUnauthenticated() {
  authUser.hidden = true;
  setAppInert(true);
  authGate.classList.remove("is-hidden");
  setGateState("unauthenticated");
  isAuthenticated = false;
  currentAccount = null;
  disarmBackGuard();
}

async function init() {
  if (e2eAuthBypassed) {
    showAuthenticated({ name: "E2E Test User", username: "e2e-test@local" });
    return;
  }
  if (!msalConfigured) {
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
      showUnauthenticated();
      showBlockedModal(error.message);
      return;
    }
    console.error("[auth] session restore failed:", error);
    setGateState("error");
    authGate.classList.remove("is-hidden");
  }
}

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
  hideBlockedModal();
});

authRetryBtn.addEventListener("click", () => {
  setGateState("loading");
  init();
});

logoutBtn.addEventListener("click", () => {
  const displayName = currentAccount?.name || "Signed in";
  const email = getAccountEmail(currentAccount) || "";
  if (displayName || email) {
    authLogoutUserName.textContent = displayName;
    authLogoutUserEmail.textContent = email;
    authLogoutUserAvatar.textContent = initialsFor(displayName);
    authLogoutUserInfo.hidden = false;
  } else {
    authLogoutUserInfo.hidden = true;
  }
  authLogoutConfirmModal.hidden = false;
});

authLogoutCancelBtn.addEventListener("click", () => {
  authLogoutConfirmModal.hidden = true;
});

authLogoutConfirmBtn.addEventListener("click", async () => {
  authLogoutConfirmBtn.disabled = true;
  authLogoutCancelBtn.disabled = true;
  try {
    await logout();
  } catch (error) {
    console.error("[auth] logout failed:", error);
    window.location.reload();
  } finally {
    authLogoutConfirmBtn.disabled = false;
    authLogoutCancelBtn.disabled = false;
    authLogoutConfirmModal.hidden = true;
  }
});

init();
