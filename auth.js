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

    // ⚠️ Attend la fin d’une redirection éventuelle
    const redirectResult = await msalInstance.handleRedirectPromise().catch(err => {
      console.warn("Erreur handleRedirectPromise:", err);
      return null;
    });

    let account = null;
    let tokenResponse = null;

    if (redirectResult && redirectResult.account) {
      console.log("Retour de redirection OK");
      account = redirectResult.account;
      msalInstance.setActiveAccount(account);
      tokenResponse = redirectResult;
    } else {
      const accounts = msalInstance.getAllAccounts();
      if (accounts.length > 0) {
        account = accounts[0];
        msalInstance.setActiveAccount(account);
      }
    }

    if (!tokenResponse) {
      if (account) {
        try {
          console.log("Tentative silencieuse...");
          tokenResponse = await msalInstance.acquireTokenSilent({
            scopes: SCOPES,
            account
          });
        } catch (e) {
          console.log("Token silencieux échoué → redirection");
          await msalInstance.acquireTokenRedirect({ scopes: SCOPES });
          return; // stop, car la redirection va s'effectuer
        }
      } else {
        console.log("Aucun compte → redirection login");
        await msalInstance.loginRedirect({ scopes: SCOPES });
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