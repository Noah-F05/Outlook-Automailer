  Office.onReady(async () => {
    console.log("Office prêt, initialisation MSAL...");

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
    console.log("MSAL initialisé");

    try {
      const redirectResult = await msalInstance.handleRedirectPromise();
      let tokenResponse = redirectResult;
      const accounts = msalInstance.getAllAccounts();

      if (!tokenResponse) {
        if (accounts.length > 0) {
          try {
            tokenResponse = await msalInstance.acquireTokenSilent({
              scopes: SCOPES,
              account: accounts[0]
            });
          } catch (err) {
            console.log("Token silencieux échoué → redirection");
            Office.context.ui.messageParent(JSON.stringify({ status: "redirecting" }));
            await msalInstance.acquireTokenRedirect({ scopes: SCOPES, account: accounts[0] });
            return;
          }
        } else {
          console.log("Aucun compte → redirection login");
          Office.context.ui.messageParent(JSON.stringify({ status: "redirecting" }));
          await msalInstance.loginRedirect({ scopes: SCOPES });
          return;
        }
      }

      const accessToken = tokenResponse?.accessToken;
      if (!accessToken) throw new Error("Impossible de récupérer le token");

      console.log("✅ Token acquis, renvoi vers Outlook add-in");
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