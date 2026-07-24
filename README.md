# EAsync Calendrier — Add-in Outlook

Agent IA de gestion de calendrier, intégré directement dans Outlook (task pane),
pour ton calendrier et ceux que des collègues t'ont déjà partagés.

Ce projet est livré **prêt à héberger**, sans étape de build de ton côté (pas
besoin de Node/npm sur ton poste).

## État actuel

- ✅ Manifest + task pane fonctionnels en **mode mock** (données simulées)
- ⏳ Connexion à Microsoft Graph : en attente du client ID/tenant ID de l'IT
- ⏳ Compréhension du langage naturel : en attente d'un petit backend (voir plus bas)

## 1. Héberger les fichiers sur GitHub Pages

1. Crée un repo GitHub (public ou privé, peu importe) — par exemple `eassist-addin`.
2. Mets-y tous les fichiers de ce dossier (`manifest.xml`, `taskpane.html`,
   `taskpane.css`, `taskpane.js`, `auth.js`, `graph.js`, `config.js`, `assets/`).
3. Dans le repo GitHub : **Settings > Pages > Source** → sélectionne la branche
   `main` et le dossier `/ (root)`. Sauvegarde.
4. GitHub te donne une URL du type `https://TON-USERNAME.github.io/eassist-addin/`.
5. Remplace **toutes** les occurrences de `VOTRE-USERNAME` dans `manifest.xml`
   et `config.js` par cette URL réelle (recherche/remplace).

## 2. Sideloader l'add-in dans Outlook

- Génère un GUID unique pour `<Id>` dans `manifest.xml` (PowerShell :
  `[guid]::NewGuid()`, ou un générateur en ligne).
- Dans Outlook sur le web : **Paramètres (roue crantée) > Gérer les compléments
  > Mes compléments > Ajouter un complément personnalisé > Ajouter depuis un
  fichier** → sélectionne ton `manifest.xml`.
- Si ça échoue avec "Sideloading rejected by Exchange", c'est un réglage tenant
  (`AppsForOfficeEnabled` ou politique de rôle Exchange) — à faire lever par
  l'IT, pas un souci côté toi.

## 3. Ce qu'il faut demander à l'IT (ticket unique, à regrouper avec la demande
   Graph déjà en cours pour EAsync)

- Créer un **App Registration** Entra ID (plateforme "Single-page application",
  redirect URI = l'URL GitHub Pages de `taskpane.html`)
- Consentement admin sur : `User.Read`, `Calendars.ReadWrite`,
  `Calendars.Read.Shared`, `Calendars.ReadWrite.Shared`
- Vérifier `AppsForOfficeEnabled` + politique de rôle Exchange pour le sideloading

Une fois reçus, remplace `clientId` et `tenantId` dans `config.js`, et passe
`mockMode` à `false`.

## 4. Point d'attention : la compréhension du langage naturel

Le fichier `taskpane.js` contient un stub très basique (`stubParsePrompt`) à la
place d'un vrai appel à un LLM. **Ce fichier tourne dans le navigateur** — on ne
peut pas y mettre une clé API Claude en clair, elle serait visible par
n'importe qui via l'inspecteur du navigateur.

Il faut un petit backend intermédiaire (proxy) qui :
1. reçoit le texte de la demande depuis le task pane,
2. appelle l'API Claude côté serveur (clé API gardée secrète côté serveur),
3. renvoie une liste d'actions structurées (JSON) au task pane.

Options légères, sans gérer de serveur en continu : Azure Function (cohérent
avec votre écosystème Microsoft), Cloudflare Worker, ou AWS Lambda. On peut
coder ce proxy ensemble dès que tu sais lequel de ces trois est le plus simple
à faire héberger côté IT.

## Structure des fichiers

```
eassist-addin/
├── manifest.xml      # déclaration de l'add-in pour Outlook
├── taskpane.html      # UI
├── taskpane.css        # styles
├── taskpane.js          # logique UI + orchestration
├── auth.js               # connexion MSAL (déléguée)
├── graph.js               # appels Microsoft Graph (calendriers/événements)
├── config.js               # clientId/tenantId/scopes — à compléter
└── assets/
    ├── icon-16.png
    ├── icon-32.png
    └── icon-80.png
```
