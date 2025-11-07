(async () => { 
  console.log("Initialisation MSAL...");

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
  console.log("✅ MSAL initialisé");

  const sendMessage = (payload) => {
    try {
      // 🔹 Stocke aussi le message dans localStorage (canal de secours)
      localStorage.setItem("automailer_auth", JSON.stringify(payload));

      if (window.Office && Office.context && Office.context.ui) {
        Office.context.ui.messageParent(JSON.stringify(payload));
      } else if (window.opener) {
        window.opener.postMessage(payload, "*");
      } else {
        console.warn("⚠️ Aucun canal de communication trouvé (ni Office, ni opener)");
      }
    } catch (err) {
      console.error("Erreur envoi message :", err);
    }
  };

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
        } catch {
          console.log("Token silencieux échoué → redirection");
          sendMessage({ status: "redirecting" });
          await msalInstance.acquireTokenRedirect({ scopes: SCOPES, account: accounts[0] });
          return;
        }
      } else {
        console.log("👤 Aucun compte → redirection login");
        sendMessage({ status: "redirecting" });
        await msalInstance.loginRedirect({ scopes: SCOPES });
        return;
      }
    }

    const accessToken = tokenResponse?.accessToken;
    if (!accessToken) throw new Error("Impossible de récupérer le token");

    console.log("✅ Token acquis !");
    sendMessage({ accessToken });
  } catch (err) {
    console.error("Auth error:", err);
    sendMessage({ error: err.message || String(err) });
  }
})();
 