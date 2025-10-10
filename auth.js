import { PublicClientApplication } from "https://cdn.jsdelivr.net/npm/@azure/msal-browser@3.24.0/+esm";

console.log("auth.js chargé");

const CLIENT_ID = "666aff85-bd89-449e-8fd7-0dbe41ed5f69";
const SCOPES = ["User.Read", "Mail.Send"];
const TENANT_ID = "2b20a610-d1f0-4c32-9cf4-0270d570bc9a";

const msalConfig = {
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    redirectUri: "https://noah-f05.github.io/Outlook-Automailer/auth.html"
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: true
  }
};

const msalInstance = new PublicClientApplication(msalConfig);

Office.onReady(async () => {
  console.log("Office prêt, initialisation MSAL...");
  try {
    await msalInstance.initialize();
    console.log("MSAL initialisé");

    const redirectResult = await msalInstance.handleRedirectPromise();
    let tokenResponse = redirectResult;
    const accounts = msalInstance.getAllAccounts();

    if (!tokenResponse) {
      if (accounts.length > 0) {
        console.log("Compte trouvé, tentative silencieuse...");
        try {
          tokenResponse = await msalInstance.acquireTokenSilent({
            scopes: SCOPES,
            account: accounts[0]
          });
        } catch (e) {
          console.log("Token silencieux échoué, redirection...");
          msalInstance.acquireTokenRedirect({ scopes: SCOPES, account: accounts[0] });
          return;
        }
      } else {
        console.log("Aucun compte, redirection login...");
        msalInstance.loginRedirect({ scopes: SCOPES });
        return;
      }
    }

    const accessToken = tokenResponse?.accessToken;
    if (!accessToken) throw new Error("Impossible de récupérer le token");

    console.log("✅ Token obtenu, envoi à Outlook");
    Office.context.ui.messageParent(JSON.stringify({ accessToken }));
  } catch (err) {
    console.error("Auth error:", err);
    try {
      Office.context.ui.messageParent(JSON.stringify({
        error: err.message || String(err)
      }));
    } catch {}
  }
});