// All values here are PUBLIC and safe to ship in client code.
// A SPA client id is not a secret (the security comes from PKCE + the
// redirect URIs you register in Azure). There is NO client secret anywhere
// in this app, by design.

// Client id comes from a Vercel/Vite env var if present, otherwise the
// hardcoded fallback below. Set VITE_MS_CLIENT_ID in Vercel → Project →
// Settings → Environment Variables, OR just paste your id into the fallback.
const FALLBACK_CLIENT_ID = "PASTE_YOUR_CLIENT_ID_HERE";

export const CLIENT_ID =
  import.meta.env.VITE_MS_CLIENT_ID || FALLBACK_CLIENT_ID;

export const msalConfig = {
  auth: {
    clientId: CLIENT_ID,
    // "consumers" = personal Microsoft accounts only (msn.com, outlook.com,
    // hotmail.com, live.com). This matches the account type you register in
    // Azure ("Personal Microsoft accounts only").
    authority: "https://login.microsoftonline.com/consumers",
    // Auto-matches whatever domain the app is served from (localhost in dev,
    // your *.vercel.app domain in prod). Register BOTH in Azure.
    redirectUri: window.location.origin,
    postLogoutRedirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false,
  },
};

// Delegated Graph permissions we ask the signed-in user to consent to.
// Mail.ReadWrite is what lets us move messages between folders.
export const loginRequest = {
  scopes: ["User.Read", "Mail.ReadWrite"],
};

export const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
export const DELETE_REVIEW_FOLDER = "Delete Review";

// Safety cap: stop paging after this many matched messages in one run.
export const MAX_MESSAGES = 5000;
