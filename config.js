// ============================================================
// CONFIG — à compléter dès que l'IT a créé l'App Registration
// ============================================================
//
// Où trouver ces valeurs dans le portail Azure/Entra :
// Entra ID > App registrations > (ton app) > Overview
//   - clientId    = "Application (client) ID"
//   - tenantId    = "Directory (tenant) ID"
// Entra ID > App registrations > (ton app) > Authentication
//   - ajouter une plateforme "Single-page application"
//   - redirectUri = l'URL GitHub Pages de taskpane.html (celle ci-dessous)
//
const CONFIG = {
  auth: {
    clientId: "REMPLACER_PAR_CLIENT_ID_FOURNI_PAR_IT",
    tenantId: "REMPLACER_PAR_TENANT_ID_FOURNI_PAR_IT",
    redirectUri: "https://aurelaye1111.github.io/eassist-addin/taskpane.html",
  },

  // Scopes délégués Graph nécessaires (doivent correspondre à ceux
  // consentis côté Azure — cf. ticket IT)
  graphScopes: [
    "User.Read",
    "Calendars.ReadWrite",
    "Calendars.Read.Shared",
    "Calendars.ReadWrite.Shared",
  ],

  // Tant que l'App Registration n'est pas prête côté IT, on travaille
  // en mode mock : aucune vraie donnée n'est lue/écrite, tout est simulé
  // pour qu'on puisse itérer sur l'UX sans attendre.
  mockMode: true,
};
