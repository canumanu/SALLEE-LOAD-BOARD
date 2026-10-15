// Sallee Load Board — site configuration.
//
// Edit this file directly on GitHub (or locally) to change sign-in or
// SharePoint settings — it is NOT part of the bundle, so changes here go
// live immediately without rebuilding bundle.js.
//
// DEMO MODE: set clientId to "" to run the board with no sign-in at all —
// spreadsheet orders only, no "+ New Order" button. Useful while the
// SharePoint/Azure AD setup in LOAD_ORDERS_SETUP.md is still in progress;
// the board works fine this way, it just doesn't have the New Order
// feature yet.
window.SALLEE_LOAD_BOARD_CONFIG = {
  // Reuses the existing "Sallee Dispatch Board Sign-in" Azure AD app
  // registration (same app the dispatch board signs in with) — this
  // board's two GitHub Pages URLs just need to be added to it as extra
  // redirect URIs. See LOAD_ORDERS_SETUP.md step 4.
  clientId: "3530a45a-659e-4d10-95fd-6f32b857e702",
  tenantId: "428e82a0-5a1b-41e6-a24c-ca8ea70db05a",
  redirectUri: "https://canumanu.github.io/SALLEE-LOAD-BOARD/",
  postLogoutRedirectUri: "https://canumanu.github.io/SALLEE-LOAD-BOARD/",

  // SharePoint site + list that orders entered through "+ New Order" are
  // saved to (separate from the ORDERS_LIST_2026.xlsx spreadsheet sync).
  // Lives on the Sallee Dispatch Board site since that's where the sign-in
  // app already has access — the Load Orders list itself has its own,
  // separate permissions (see LOAD_ORDERS_SETUP.md step 3).
  spHostname: "salleehorsevans.sharepoint.com",
  spSitePath: "/sites/SALLEEDISPATCHBOARD",
  listName: "Load Orders",

  // How often (ms) the board re-polls the Load Orders list while signed in.
  pollIntervalMs: 30000,
};
