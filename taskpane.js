(() => {
  const AUTH_PAGE = "https://noah-f05.github.io/Outlook-Automailer/auth.html";
  let accessToken = null;

  function log(msg) {
    console.log("[Add-in]", msg);
  }

  async function openAuthDialog() {
    return new Promise((resolve, reject) => {
      log("Ouverture du dialogue d'auth...");
      Office.context.ui.displayDialogAsync(
        AUTH_PAGE,
        { height: 60, width: 40, displayInIframe: false },
        (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            return reject(new Error("Impossible d'ouvrir la fenêtre d'auth"));
          }

          const dialog = asyncResult.value;

          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
            try {
              const payload = JSON.parse(arg.message);

              // 🟡 Nouvelle gestion : si auth.html est en redirection, on ignore
              if (payload.status === "redirecting") {
                log("L'utilisateur est redirigé vers Microsoft Login...");
                return;
              }

              if (payload.accessToken) {
                resolve(payload.accessToken);
              } else {
                reject(new Error(payload.error || "Message inattendu depuis auth.html"));
              }
            } catch (e) {
              reject(e);
            } finally {
              try { dialog.close(); } catch {}
            }
          });

          dialog.addEventHandler(Office.EventType.DialogEventReceived, (event) => {
            reject(new Error("Fenêtre d'auth fermée/annulée: " + JSON.stringify(event)));
            try { dialog.close(); } catch {}
          });
        }
      );
    });
  }

  async function getRecipientsAsync() {
  const getField = (fieldName) => new Promise((resolve) => {
    Office.context.mailbox.item[fieldName].getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const emails = (result.value || [])
          .map(r => (typeof r === "string" ? r : (r.emailAddress || r.address)))
          .filter(Boolean);
        resolve(emails);
      } else {
        resolve([]); // on ignore les erreurs partielles
      }
    });
  });

  // Récupère les 3 types de destinataires
  const [to, cc, bcc] = await Promise.all([
    getField("to"),
    getField("cc"),
    getField("bcc")
  ]);

  // Fusionne tout en supprimant les doublons
  const allRecipients = [...new Set([...to, ...cc, ...bcc])];

  console.log(`📬 Destinataires trouvés : ${allRecipients.join(", ")}`);
  return allRecipients;
}


  async function getSubjectAsync() {
    return new Promise((resolve, reject) => {
      Office.context.mailbox.item.subject.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value || "(Sans sujet)");
        } else {
          reject(result.error || new Error("Impossible de lire le sujet du mail"));
        }
      });
    });
  }

  async function getBodyHtmlAsync() {
    return new Promise((resolve, reject) => {
      Office.context.mailbox.item.body.getAsync("html", (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value || "");
        } else {
          reject(result.error || new Error("Impossible de lire le corps"));
        }
      });
    });
  }

  async function getAttachmentsFromDraft(itemId, token) {
    if (!itemId) return [];

    const restId = Office.context.mailbox.convertToRestId(
      itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );

    const res = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${restId}/attachments`, {
      method: "GET",
      headers: {
        "Authorization": `Bearer ${token}`,
        "Content-Type": "application/json"
      }
    });

    if (!res.ok) {
      const txt = await res.text().catch(() => "");
      throw new Error(`Graph error ${res.status}: ${txt}`);
    }

    const data = await res.json();
    return (data.value || [])
      .filter(att => att["@odata.type"] === "#microsoft.graph.fileAttachment")
      .map(att => ({
        name: att.name,
        contentBytes: att.contentBytes,
        contentType: att.contentType || "application/octet-stream"
      }));
  }

  async function sendEmail(token, to, subject, bodyHtml, attachments = [], cc = [], bcc = []) {
    const mail = {
      message: {
        subject: subject || "(Sans sujet)",
        body: { contentType: "HTML", content: bodyHtml || "" },
        toRecipients: to.map(addr => ({ emailAddress: { address: addr } })),
        attachments: attachments.map(att => ({
          "@odata.type": "#microsoft.graph.fileAttachment",
          name: att.name,
          contentBytes: att.contentBytes,
          contentType: att.contentType || "application/octet-stream"
        }))
      }
    };

    const res = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${token}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify(mail)
    });

    if (!res.ok) {
      const txt = await res.text().catch(() => "");
      throw new Error(`Graph error ${res.status}: ${txt}`);
    }
  }

  async function saveDraftIfNeeded() {
    return new Promise((resolve, reject) => {
      Office.context.mailbox.item.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(result.error || new Error("Impossible de sauvegarder le draft"));
        }
      });
    });
  }

  async function deleteDraft(itemId, accessToken) {
    if (!itemId) return;
    const restId = Office.context.mailbox.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0);

    const res = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${restId}`, {
      method: "DELETE",
      headers: { "Authorization": `Bearer ${accessToken}` }
    });

    if (!res.ok) {
      const txt = await res.text().catch(() => "");
      console.error("Erreur suppression draft:", txt);
    } else {
      console.log("Draft supprimé avec succès");
    }
  }

  async function run() {
    try {
      log("Démarrage de l'action...");
      let draftId = null;
      try {
        draftId = await saveDraftIfNeeded();
        log("Draft sauvegardé, ID = " + draftId);
      } catch (err) {
        console.warn("Impossible de sauvegarder le draft :", err);
      }

      if (!accessToken) {
        log("Pas de token — ouverture auth...");
        accessToken = await openAuthDialog();
      }

      if (!accessToken) throw new Error("Token introuvable après auth");

      const recipients = await getRecipientsAsync();
      if (!recipients || recipients.length === 0) {
        Office.context.mailbox.item.notificationMessages.replaceAsync("noRecipients", {
          type: "errorMessage",
          message: "⚠️ Aucun destinataire trouvé."
        });
        return;
      }
      const subject = await getSubjectAsync();
      const bodyHtml = await getBodyHtmlAsync();

      let attachments = [];
      if (draftId) {
        try {
          attachments = await getAttachmentsFromDraft(draftId, accessToken);
          log(`📎 ${attachments.length} pièce(s) jointe(s) récupérée(s)`);
        } catch (err) {
          console.warn("Impossible de récupérer les pièces jointes :", err);
        }
      }

      let sent = 0;
      for (const to of recipients) {
        try {
          log(`Envoi à ${to}...`);
          await sendEmail(accessToken, to, subject, bodyHtml, attachments);
          sent++;
        } catch (err) {
          console.error("Erreur envoi:", err);
        }
      }

      if (draftId && sent === recipients.length) {
        await deleteDraft(draftId, accessToken);
      }

      Office.context.mailbox.item.notificationMessages.replaceAsync("successMsg", {
        type: "informationalMessage",
        message: `✅ ${sent} email(s) envoyés individuellement`,
        icon: "icon16",
        persistent: false
      });

      log("Terminé.");
    } catch (err) {
      console.error("Erreur run:", err);
      Office.context.mailbox.item.notificationMessages.replaceAsync("errorMsg", {
        type: "errorMessage",
        message: "❌ Une erreur est survenue. Voir console."
      });
    }
  }

  Office.onReady(() => {
    Office.actions.associate("run", run);
  });
})();
