// ============================================================
// AUTH — delegated sign-in via MSAL.js (auth code flow + PKCE)
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
      // sessionStorage: avoids keeping a token around if someone else
      // uses the same machine later
      cacheLocation: "sessionStorage",
    },
  });

  return msalInstance;
}

async function signIn() {
  if (CONFIG.mockMode) {
    activeAccount = { name: "Aurelie (demo mode)", username: "demo@raileurope.com" };
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
  if (!account) throw new Error("No active session — please sign in first.");

  try {
    const result = await instance.acquireTokenSilent({
      scopes: CONFIG.graphScopes,
      account,
    });
    return result.accessToken;
  } catch (err) {
    // Silent token acquisition failed (expired, missing consent...):
    // retry via popup, which will re-prompt for consent if needed.
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
