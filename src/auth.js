import {
  PublicClientApplication,
  InteractionRequiredAuthError,
} from "@azure/msal-browser";
import { msalConfig, loginRequest } from "./config.js";

export const msalInstance = new PublicClientApplication(msalConfig);

let initialized = false;

// Must be called once on page load, BEFORE any other MSAL call (MSAL v3/v4
// requires an explicit initialize). Also processes the redirect that brings
// the user back after signing in.
export async function initAuth() {
  if (!initialized) {
    await msalInstance.initialize();
    initialized = true;
  }

  const result = await msalInstance.handleRedirectPromise();
  if (result && result.account) {
    msalInstance.setActiveAccount(result.account);
  } else if (!msalInstance.getActiveAccount()) {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      msalInstance.setActiveAccount(accounts[0]);
    }
  }
  return msalInstance.getActiveAccount();
}

export function getActiveAccount() {
  return msalInstance.getActiveAccount();
}

// prompt: "select_account" lets you pick which personal account to use,
// which matters if f14stalker@msn.com and stalkersignup@outlook.com are
// separate accounts rather than aliases on one account.
export function login() {
  return msalInstance.loginRedirect({
    ...loginRequest,
    prompt: "select_account",
  });
}

export function logout() {
  return msalInstance.logoutRedirect({
    account: msalInstance.getActiveAccount(),
  });
}

// Returns a fresh Graph access token, silently refreshing when possible.
// If interaction is required it redirects (and resolves to null because the
// page is navigating away).
export async function getToken() {
  const account = msalInstance.getActiveAccount();
  if (!account) throw new Error("Not signed in.");
  try {
    const res = await msalInstance.acquireTokenSilent({
      ...loginRequest,
      account,
    });
    return res.accessToken;
  } catch (err) {
    if (err instanceof InteractionRequiredAuthError) {
      await msalInstance.acquireTokenRedirect(loginRequest);
      return null;
    }
    throw err;
  }
}
