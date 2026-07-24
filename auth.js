// ============================================================
// AUTH — connexion déléguée via MSAL.js (auth code flow + PKCE)
// ============================================================

let msalInstance = null;
let activeAccount = null;

function getMsalInstance() {
  if (msalInstance) return msalInstance;

  msalInstance = new msal.PublicClientApplication({
    auth: {
      clientId: CONFIG.auth.clientId,
      authority: `https://login.microsoftonline.com/${CONFIG.auth.tenantId}`,
      redirectUri: CONFIG.auth.redirectUri,
    },
    cache: {
      // sessionStorage : évite de garder un token si quelqu'un d'autre
      // reprend le même poste plus tard
      cacheLocation: "sessionStorage",
    },
  });

  return msalInstance;
}

async function signIn() {
  if (CONFIG.mockMode) {
    activeAccount = { name: "Aurélie (mode démo)", username: "demo@raileurope.com" };
    return activeAccount;
  }

  const instance = getMsalInstance();
  const result = await instance.loginPopup({ scopes: CONFIG.graphScopes });
  instance.setActiveAccount(result.account);
  activeAccount = result.account;
  return activeAccount;
}

async function getAccessToken() {
  if (CONFIG.mockMode) return "mock-token";

  const instance = getMsalInstance();
  const account = instance.getActiveAccount();
  if (!account) throw new Error("Pas de session active — connecte-toi d'abord.");

  try {
    const result = await instance.acquireTokenSilent({
      scopes: CONFIG.graphScopes,
      account,
    });
    return result.accessToken;
  } catch (err) {
    // Le token silencieux a échoué (expiré, consentement manquant...) :
    // on retente via popup, qui redemandera un consentement si besoin.
    const result = await instance.acquireTokenPopup({ scopes: CONFIG.graphScopes });
    return result.accessToken;
  }
}

function isSignedIn() {
  return !!activeAccount;
}

function getActiveAccount() {
  return activeAccount;
}
