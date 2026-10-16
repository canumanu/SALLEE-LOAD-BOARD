// MSAL + Microsoft Graph plumbing for the Load Board's sign-in.
//
// Same pattern as the Sallee Dispatch Board (see that repo's index.html,
// the "AUTH / SHAREPOINT SYNC" section) — reused here against the exact
// same "Sallee Dispatch Board Sign-in" Azure AD app registration, just
// with this board's URL added as an extra redirect URI. Kept as plain
// functions + module-level state (no React) so sharepoint.js and
// OrderForm.jsx can both use it without prop-drilling a client instance
// around.

const cfg = (typeof window !== "undefined" && window.SALLEE_LOAD_BOARD_CONFIG) || {};

// Demo mode: no clientId configured in config.js -> sign-in is skipped
// entirely and the board runs spreadsheet-orders-only. This is the state
// the board was shipped live in before Azure AD / SharePoint setup was
// finished, and remains a safe fallback if that setup ever needs undoing.
export const DEMO_MODE = !cfg.clientId;

const MSAL_CONFIG = {
  auth: {
    clientId: cfg.clientId,
    authority: `https://login.microsoftonline.com/${cfg.tenantId}`,
    redirectUri: cfg.redirectUri,
    postLogoutRedirectUri: cfg.postLogoutRedirectUri,
  },
  cache: {
    // localStorage (not sessionStorage) so people stay signed in across
    // reloads and new tabs, not just within one tab session.
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false,
  },
};
const GRAPH_SCOPES = ["Sites.ReadWrite.All"];
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const SITE_ID_CACHE_KEY = "sallee-load-board:sp-site-id";

let msalInstance = null;
let currentAccount = null;
let siteId = null;
const listIdCache = {};

export function getAccount() {
  return currentAccount;
}

// Call once on startup. Resolves to the signed-in account if one already
// exists (returning from a redirect, or a cached session), or null.
export async function initAuth() {
  if (DEMO_MODE) return null;
  if (typeof msal === "undefined") {
    console.error("MSAL library failed to load from CDN — sign-in unavailable this load");
    return null;
  }
  msalInstance = new msal.PublicClientApplication(MSAL_CONFIG);
  await msalInstance.initialize();
  const response = await msalInstance.handleRedirectPromise();
  if (response && response.account) {
    currentAccount = response.account;
  } else {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) currentAccount = accounts[0];
  }
  if (currentAccount) msalInstance.setActiveAccount(currentAccount);
  return currentAccount;
}

export function signIn() {
  if (!msalInstance) {
    throw new Error("Sign-in is unavailable right now (auth library failed to load). Try refreshing the page.");
  }
  msalInstance.loginRedirect({ scopes: GRAPH_SCOPES });
}

// Local-only sign-out: clears this board's own session in this browser and
// reloads to the sign-in screen. Deliberately NOT msalInstance.logoutRedirect()
// -- that performs a full Microsoft 365 sign-out (ending Outlook/Teams/etc
// sessions too), which is outside what "sign out of the load board" should do.
//
// Uses the token cache's own removeAccount() rather than pattern-matching
// localStorage keys -- MSAL's actual account/token cache entries are keyed
// by `homeAccountId-environment-realm` and friends, not a "msal." prefix
// and not guaranteed to contain the clientId, so the old key-guessing
// approach silently left the account entry behind and the board signed
// itself right back in on reload.
export async function signOut() {
  try {
    sessionStorage.removeItem(SITE_ID_CACHE_KEY);
    if (msalInstance) {
      const cache = msalInstance.getTokenCache();
      const accounts = msalInstance.getAllAccounts();
      for (const account of accounts) {
        await cache.removeAccount(account);
      }
    }
  } catch {
    // ignore storage/cache errors (private browsing etc.) -- the reload
    // below is still a safe fallback even if cache clearing partially fails.
  }
  currentAccount = null;
  location.reload();
}

async function getAccessToken() {
  const account = currentAccount || (msalInstance && msalInstance.getActiveAccount());
  if (!msalInstance || !account) throw new Error("Not signed in");
  try {
    const result = await msalInstance.acquireTokenSilent({ scopes: GRAPH_SCOPES, account });
    return result.accessToken;
  } catch (err) {
    if (typeof msal !== "undefined" && err instanceof msal.InteractionRequiredAuthError) {
      await msalInstance.acquireTokenRedirect({ scopes: GRAPH_SCOPES, account });
      return null; // navigates away; nothing after this point runs
    }
    throw err;
  }
}

export async function graphFetch(pathOrUrl, options) {
  const token = await getAccessToken();
  if (!token) throw new Error("No access token");
  const url = pathOrUrl.indexOf("http") === 0 ? pathOrUrl : GRAPH_BASE + pathOrUrl;
  // Only send Content-Type when there's an actual JSON body (POST/PATCH) --
  // a bodyless GET carrying Content-Type can come back as a misleading 403
  // from some SharePoint-backed Graph endpoints (bit the dispatch board in
  // testing; carried the same caution over here).
  const hasBody = !!(options && options.body != null);
  const headers = {
    Authorization: "Bearer " + token,
    ...(hasBody ? { "Content-Type": "application/json" } : {}),
    ...((options && options.headers) || {}),
  };
  const res = await fetch(url, { ...options, headers });
  if (!res.ok) {
    let bodyText = "";
    try {
      bodyText = await res.text();
    } catch {
      // ignore
    }
    const err = new Error(`Graph API ${res.status}: ${bodyText}`);
    err.status = res.status;
    throw err;
  }
  if (res.status === 204) return null;
  return res.json();
}

export async function resolveSite() {
  if (siteId) return siteId;
  const cached = sessionStorage.getItem(SITE_ID_CACHE_KEY);
  if (cached) {
    siteId = cached;
    return siteId;
  }
  const site = await graphFetch(`/sites/${cfg.spHostname}:${cfg.spSitePath}`);
  siteId = site.id;
  sessionStorage.setItem(SITE_ID_CACHE_KEY, siteId);
  return siteId;
}

// Finds an existing list by display name. Deliberately does NOT try to
// create one if it's missing -- list/column creation via Graph is blocked
// on this tenant (same finding as the dispatch board), and the Load
// Orders list is meant to be created once by hand from
// setup/Load_Orders_list_template.xlsx (see LOAD_ORDERS_SETUP.md). A
// missing list means setup isn't finished, not something to paper over.
export async function findList(name) {
  if (listIdCache[name]) return listIdCache[name];
  const sid = await resolveSite();
  const allLists = await graphFetch(`/sites/${sid}/lists`);
  const match = (allLists.value || []).find((l) => l.displayName === name);
  if (!match) {
    const err = new Error(`SharePoint list "${name}" wasn't found on this site yet`);
    err.code = "LIST_NOT_FOUND";
    throw err;
  }
  listIdCache[name] = match.id;
  return match.id;
}
