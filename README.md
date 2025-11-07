# 📧 Outlook Automailer Add-in

Ce projet est un complément Outlook (Office Add-in) développé pour **<Nom de l’entreprise>**.  
Il permet d’envoyer automatiquement un **mail individuel** à chaque destinataire (TO, CC, CCI) d’un brouillon, tout en **conservant le contenu, la mise en forme, la signature et les pièces jointes** du message original.

L’objectif est d’éviter les envois groupés et de garantir la confidentialité entre destinataires tout en facilitant la gestion des envois multiples.

---

## ⚙️ Fonctionnalités principales

- 🔐 Authentification sécurisée via **Microsoft 365 (MSAL.js + OAuth2)**
- 📩 Lecture complète du brouillon Outlook (sujet, corps, pièces jointes, images inline)
- ✉️ Envoi **individuel** via **Microsoft Graph API** (`/me/sendMail`)
- 🧾 Suppression automatique du brouillon après envoi
- 🖼️ Gestion correcte des images intégrées à la signature
- 🌐 Compatible avec **Outlook Web** et **Outlook Desktop (Windows / Edge / Chrome)**

---

## 🧱 Structure du projet

Outlook-Automailer/
│
├── manifest.xml # Déclaration du complément Outlook
├── taskpane.html # Interface utilisateur (panneau latéral)
├── taskpane.js # Logique principale (lecture mail + envoi Graph)
├── auth.html # Page d’authentification Microsoft
├── auth.js # Gestion de l’authentification MSAL
├── assets/ # Dossier contenant les icônes, logos, images
└── README.md # Documentation du projet


---

## ☁️ Hébergement et infrastructure

Le complément est hébergé sur **Azure Static Web Apps**, sous le tenant de l’entreprise.

### 🔗 URLs principales
| Élément | URL |
|----------|-----|
| Taskpane | `https://outlook-automailer.<entreprise>.azurestaticapps.net/taskpane.html` |
| Auth page | `https://outlook-automailer.<entreprise>.azurestaticapps.net/auth.html` |
| Redirect URI (Azure AD) | même URL que `auth.html` |

### 🚀 Déploiement automatique
Un workflow **GitHub Actions** déploie automatiquement le site sur Azure à chaque *push* sur la branche `main`.

---

## 🔐 Authentification Microsoft (MSAL)

Le complément utilise **MSAL.js** pour gérer l’authentification et les permissions Microsoft Graph.

### Configuration MSAL
const msalConfig = {
  auth: {
    clientId: "<CLIENT_ID>",
    authority: "https://login.microsoftonline.com/<TENANT_ID>",
    redirectUri: "https://outlook-automailer.<entreprise>.azurestaticapps.net/auth.html"
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: true
  }
};
