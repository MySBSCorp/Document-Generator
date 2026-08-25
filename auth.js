import {
  PublicClientApplication,
  InteractionRequiredAuthError,
  ServerError,
  EventType,
} from "./vendor/msal-browser-4.30.0.esm.js";
import { msalConfig, loginRequest, allowedEmails } from "./authConfig.js";

export const msalInstance = new PublicClientApplication(msalConfig);

export class AccessDeniedError extends Error {
  constructor(message) {
    super(message || "This account is not authorized to use this application.");
    this.name = "AccessDeniedError";
  }
}

function getAccountEmail(account) {
  return account?.idTokenClaims?.email || account?.username || null;
}

function isEmailAllowed(account) {
  if (!allowedEmails.length) return true;
  const email = getAccountEmail(account);
  return Boolean(email && allowedEmails.includes(email.toLowerCase()));
}

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
let pendingRedirectAccount;
async function ensureInitialized() {
  if (initialized) return;
  await msalInstance.initialize();
  try {
    const result = await msalInstance.handleRedirectPromise();
    pendingRedirectAccount = result?.account ?? null;
  } catch (error) {
    initialized = true;
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

export async function restoreSession() {
  await ensureInitialized();

  if (pendingRedirectAccount) {
    const account = pendingRedirectAccount;
    pendingRedirectAccount = null;
    if (!isEmailAllowed(account)) {
      throw new AccessDeniedError("This account isn't on the allowed list for this app.");
    }
    msalInstance.setActiveAccount(account);
    return account;
  }

  const account = getActiveAccount();
  if (!account) return null;
  if (!isEmailAllowed(account)) {
    throw new AccessDeniedError("This account isn't on the allowed list for this app.");
  }
  try {
    await msalInstance.acquireTokenSilent({ scopes: loginRequest.scopes, account });
    return account;
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError) {
      return null;
    }
    throw error;
  }
}

export async function login(options = {}) {
  await ensureInitialized();
  await msalInstance.loginRedirect({ ...loginRequest, ...options });
}

export async function logout() {
  await ensureInitialized();
  const account = getActiveAccount();
  await msalInstance.logoutRedirect({ account });
}

export { getAccountEmail };
