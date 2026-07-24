// ============================================================
// CONFIG — fill in once IT has created the App Registration
// ============================================================
//
// Where to find these values in the Azure/Entra portal:
// Entra ID > App registrations > (your app) > Overview
//   - clientId    = "Application (client) ID"
//   - tenantId    = "Directory (tenant) ID"
// Entra ID > App registrations > (your app) > Authentication
//   - add a "Single-page application" platform
//   - redirectUri = the GitHub Pages URL of taskpane.html (below)
//
const CONFIG = {
  auth: {
    clientId: "REPLACE_WITH_CLIENT_ID_FROM_IT",
    tenantId: "REPLACE_WITH_TENANT_ID_FROM_IT",
    redirectUri: "https://aurelaye1111.github.io/eassist-addin/taskpane.html",
  },

  // Required delegated Graph scopes (must match the ones consented
  // to on the Azure side — see the IT ticket)
  graphScopes: [
    "User.Read",
    "Calendars.ReadWrite",
    "Calendars.Read.Shared",
    "Calendars.ReadWrite.Shared",
  ],

  // Until the App Registration is ready on IT's side, we work in
  // mock mode: no real data is read or written, everything is
  // simulated so we can iterate on the UX without waiting.
  mockMode: true,
};
