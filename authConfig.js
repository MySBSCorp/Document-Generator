const env = import.meta.env;

const clientId = env?.VITE_MSAL_CLIENT_ID;

const tenantId = env?.VITE_MSAL_TENANT_ID;

const redirectUri = env?.VITE_MSAL_REDIRECT_URI || window.location.origin;

export const allowedEmails = (env?.VITE_ALLOWED_EMAILS || "")
  .split(",")
  .map((email) => email.trim().toLowerCase())
  .filter(Boolean);

export const msalConfigured = Boolean(clientId && tenantId);

export const msalConfig = {
  auth: {
    clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    redirectUri,
    postLogoutRedirectUri: redirectUri,
    navigateToLoginRequestUrl: true,
  },
  cache: {
    cacheLocation: "sessionStorage",
  },
  system: {
    loggerOptions: {
      loggerCallback: (level, message, containsPii) => {
        if (containsPii) return;
        if (level === msalLogLevelError) console.error("[MSAL]", message);
      },
      piiLoggingEnabled: false,
    },
  },
};

const msalLogLevelError = 0;

export const loginRequest = {
  scopes: ["openid", "profile", "email"],
  prompt: "login",
};
