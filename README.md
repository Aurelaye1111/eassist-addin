# EAsync Calendar — Outlook Add-in

AI calendar agent, embedded directly in Outlook (task pane), for your
calendar and any calendars colleagues have already shared with you.

This project ships ready to host — no build step on your side needed
(no Node/npm required on your machine).

## Current status

- ✅ Manifest + task pane working in **mock mode** (simulated data)
- ⏳ Microsoft Graph connection: waiting on IT for the client ID/tenant ID
- ⏳ Natural-language understanding: waiting on a small backend (see below)

## 1. Host the files on GitHub Pages

1. Create a GitHub repo (public or private) — e.g. `eassist-addin`.
2. Put all the files from this folder in it (`manifest.xml`, `taskpane.html`,
   `taskpane.css`, `taskpane.js`, `auth.js`, `graph.js`, `config.js`, `assets/`).
3. In the repo: **Settings > Pages > Source** → select the `main` branch and
   `/ (root)` folder. Save.
4. GitHub gives you a URL like `https://YOUR-USERNAME.github.io/eassist-addin/`.
5. Replace every occurrence of the placeholder host in `manifest.xml` and
   `config.js` with this real URL.

## 2. Sideload the add-in in Outlook

- Generate a unique GUID for `<Id>` in `manifest.xml` (PowerShell:
  `[guid]::NewGuid()`, or an online generator).
- In Outlook on the web: **Settings (gear icon) > Manage add-ins > My add-ins
  > Add a custom add-in > Add from File** → select your `manifest.xml`.
- If it fails with "Sideloading rejected by Exchange", that's a tenant setting
  (`AppsForOfficeEnabled` or Exchange role assignment policy) — for IT, not
  something to chase on your side.

## 3. What to ask IT for (a single ticket, ideally grouped with the Graph
   request already in progress for EAsync)

- Create an **App Registration** in Entra ID ("Single-page application"
  platform, redirect URI = the GitHub Pages URL of `taskpane.html`)
- Admin consent for: `User.Read`, `Calendars.ReadWrite`,
  `Calendars.Read.Shared`, `Calendars.ReadWrite.Shared`
- Check `AppsForOfficeEnabled` + Exchange role assignment policy for sideloading

Once received, fill in `clientId` and `tenantId` in `config.js`, and set
`mockMode` to `false`.

## 4. A note on natural-language understanding

`taskpane.js` currently contains a very basic stub (`stubParsePrompt`)
instead of a real LLM call. **This file runs in the browser** — we can't put
a Claude API key in plain text here, it would be visible to anyone via the
browser inspector.

A small backend proxy is needed that:
1. receives the request text from the task pane,
2. calls the Claude API server-side (API key kept secret server-side),
3. returns a structured list of actions (JSON) back to the task pane.

Lightweight options, no server to run continuously: Azure Function
(consistent with your Microsoft ecosystem), Cloudflare Worker, or AWS Lambda.
We can build this proxy together once you know which of these three is
easiest to get IT to host.

## File structure

```
eassist-addin/
├── manifest.xml      # add-in declaration for Outlook
├── taskpane.html      # UI
├── taskpane.css        # styles
├── taskpane.js          # UI logic + orchestration
├── auth.js               # MSAL sign-in (delegated)
├── graph.js               # Microsoft Graph calls (calendars/events)
├── config.js               # clientId/tenantId/scopes — to fill in
└── assets/
    ├── icon-16.png
    ├── icon-32.png
    ├── icon-80.png
    └── icon-128.png
```
