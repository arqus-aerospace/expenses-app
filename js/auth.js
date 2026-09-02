// Microsoft sign-in (MSAL.js, redirect flow — works on phones and desktops).
import { CONFIG } from "./config.js";

const SCOPES = ["User.Read", "Sites.ReadWrite.All", "Files.ReadWrite.All"];

let pca = null;
let account = null;

function client() {
  if (!pca) {
    pca = new msal.PublicClientApplication({
      auth: {
        clientId: CONFIG.clientId,
        authority: `https://login.microsoftonline.com/${CONFIG.tenant}`,
        // GitHub Pages serves the app from a sub-path; send users back to
        // exactly where the app lives.
        redirectUri: window.location.origin + window.location.pathname,
      },
      cache: { cacheLocation: "localStorage" },
    });
  }
  return pca;
}

// Call once on startup. Resolves with the signed-in account or null.
export async function initAuth() {
  const app = client();
  await app.initialize();
  const result = await app.handleRedirectPromise();
  account = result?.account ?? app.getAllAccounts()[0] ?? null;
  if (account) app.setActiveAccount(account);
  return account;
}

export function signIn() {
  return client().loginRedirect({ scopes: SCOPES, prompt: "select_account" });
}

export function signOut() {
  sessionStorage.clear();
  return client().logoutRedirect({ account });
}

export function currentUser() {
  if (!account) return null;
  return {
    name: account.name || account.username,
    email: (account.username || "").toLowerCase(),
    // Directory the account belongs to ("tid" claim); the app refuses any
    // account outside the company tenant (see config.isCompanyAccount).
    tenantId: (account.tenantId || account.idTokenClaims?.tid || "").toLowerCase(),
  };
}

// Drop the cached account without a round-trip to Microsoft, so a refused
// sign-in doesn't come straight back on the next page load.
export function forgetAccount() {
  const app = client();
  if (account) {
    // clearCache is the v3 API; removeAccount is kept as a fallback.
    if (typeof app.clearCache === "function") app.clearCache({ account });
    else app.removeAccount?.(account);
  }
  account = null;
  sessionStorage.clear();
}

// Access token for Microsoft Graph; silently refreshed, falls back to a
// redirect when the refresh token has expired.
export async function getToken() {
  const app = client();
  try {
    const res = await app.acquireTokenSilent({ scopes: SCOPES, account });
    return res.accessToken;
  } catch (e) {
    if (e instanceof msal.InteractionRequiredAuthError) {
      await app.acquireTokenRedirect({ scopes: SCOPES, account });
      return new Promise(() => {}); // page is navigating away
    }
    throw e;
  }
}
