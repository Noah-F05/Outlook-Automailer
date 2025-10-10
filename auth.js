import * as msal from "@azure/msal-browser";

(async () => {
  const CLIENT_ID = "666aff85-bd89-449e-8fd7-0dbe41ed5f69";
  const SCOPES = ["User.Read", "Mail.Send"];

  const msalConfig = {
    auth: {
      clientId: CLIENT_ID,
      authority: "https://login.microsoftonline.com/2b20a610-d1f0-4c32-9cf4-0270d570bc9a",
      redirectUri: "https://noah-f05.github.io/Outlook-Automailer/auth.html"
    },
    cache: {
      cacheLocation: "localStorage",
      storeAuthStateInCookie: true
    }
  };

  const msalInstance = new msal.PublicClientApplication(msalConfig);
  await msalInstance.initialize();

  Office.onReady(async () => {
    try {
      const accounts = msalInstance.getAllAccounts();
      let account = accounts.length > 0 ? accounts[0] : null;

      if (!account) {
        console.log("🔐 Première connexion — ouverture du popup MSAL");
        const loginResult = await msalInstance.loginPopup({ scopes: SCOPES });
        account = loginResult.account;
      }

      console.log("✅ Compte connecté :", account?.username);

      const tokenResponse = await msalInstance.acquireTokenSilent({
        scopes: SCOPES,
        account
      }).catch(async (err) => {
        console.warn("⚠️ Silent token échoué :", err);
        return await msalInstance.acquireTokenPopup({ scopes: SCOPES });
      });

      if (!tokenResponse?.accessToken) throw new Error("Impossible d’obtenir le token");

      console.log("🎟️ Token obtenu, envoi à Outlook...");
      Office.context.ui.messageParent(JSON.stringify({ accessToken: tokenResponse.accessToken }));
      window.close();
    } catch (err) {
      console.error("❌ Auth error:", err);
      try {
        Office.context.ui.messageParent(JSON.stringify({
          error: err.message || String(err)
        }));
      } catch {}
    }
  });
})();