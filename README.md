# Sender Sweep — Outlook Sender Cleaner

A personal web app for cleaning up and manipulating my Outlook mail, operated
entirely through a Vercel-hosted page in the browser. **Moving every message
from a given sender into a "Delete Review" folder was step one.** The project
is open-ended and will grow to cover more inbox cleanup and management tasks
over time — the current feature set is a starting point, not the boundary.

Built for personal Microsoft accounts (`@msn.com`, `@outlook.com`,
`@hotmail.com`, `@live.com`).

## Where it lives

- **Repo:** `ed-cos-system/outlook-sender-cleaner` — work on `main`; pushing to
  `main` auto-deploys.
- **Live app:** <https://outlook-sender-cleaner.vercel.app/> — always open and
  test from this production domain, never a hashed per-deployment URL (the auth
  redirect only matches this one).
- **Hosting:** Vercel, auto-detects Vite, builds `main` on every push.

## How it works (and why there are no secrets)

- A Vite single-page app, vanilla JS (no framework). Files: `index.html`,
  `src/main.js` (UI), `src/auth.js` (MSAL), `src/graph.js` (Microsoft Graph
  calls), `src/config.js` (settings), `src/style.css`.
- Auth is **MSAL** (`@azure/msal-browser`) using the OAuth 2.0 **Authorization
  Code flow with PKCE** — a public client with **no client secret**. The only
  config is the app's **client id**, which is not sensitive.
- All mail access happens in the browser with a delegated token via the
  Microsoft Graph API. No server, no database, nothing stored anywhere except
  the MSAL token cache in the browser.
- Personal Microsoft accounts only (`authority: /consumers`).
- Graph scopes currently consented: **`User.Read`**, **`Mail.ReadWrite`**. A
  feature needing a broader scope means an Entra app-registration change.

## One-time setup

### 1. Get a directory + register an app

A personal Microsoft account can no longer register apps "outside a directory."
To get a directory, sign up for a **free Azure account**
(<https://azure.microsoft.com/free>) with your Microsoft account — it asks for
a card for identity verification but isn't charged, and everything this tool
uses (the app registration and Entra ID free tier) is free. Once that's done,
App registrations works:

1. Go to <https://entra.microsoft.com> (or App registrations in
   <https://portal.azure.com>).
2. **App registrations → New registration.**
3. Name: anything (e.g. `Outlook Sender Cleaner`).
4. **Supported account types:** **Personal Microsoft accounts only** (matches
   the `/consumers` authority).
5. **Redirect URI:** platform = **Single-page application (SPA)**, value =
   `https://outlook-sender-cleaner.vercel.app` — **no trailing slash**. (Also
   add `http://localhost:5173` for local dev.)
6. **Register**, then copy the **Application (client) ID**.
7. **API permissions → Microsoft Graph → Delegated** → add **`Mail.ReadWrite`**
   and **`User.Read`**.

No client secret is needed. The directory can be owned by a different account
(e.g. `Ed_Chandler@outlook.com`) than the mailbox you sign in with, because the
app is set to "personal accounts."

### 2. Give the app your client id

In Vercel → project → **Settings → Environment Variables**, add
`VITE_MS_CLIENT_ID` = your client id (check Production), then **redeploy**.
Vite inlines env vars at build time, so a redeploy is required after changing
it. (Alternatively, paste the id into `src/config.js`.)

### 3. Deploy

The repo is connected to Vercel; it auto-detects Vite (`build` → `vite build`,
output `dist`). Push to `main` and it deploys. The live URL must exactly match
the SPA redirect URI registered above.

## Local development

```bash
npm install
VITE_MS_CLIENT_ID=your-client-id npm run dev   # http://localhost:5173
npm run build                                  # production build into dist/
```

## Using it

1. **Sign in with Microsoft** — pick the account that receives the mail.
2. **Sender:** paste it exactly as Outlook shows it —
   `Name <address@example.com>` — or just the bare address. The app pulls out
   the address between the `<>` and shows you the cleaned value it will match.
3. Choose **Inbox only** (default) or **All folders**.
4. **Find messages** — shows the count and a preview of subjects.
5. **Move to Delete Review** — creates the folder if needed and moves
   everything in batches.

## Accounts

`f14stalker@msn.com` and `stalkersignup@outlook.com` both land in the same
inbox. If they're aliases on one account, sign in once. If they're separate
accounts, use **Sign out** and sign in to the other to sweep it too
(sign-in uses "select account").

## Working style / gotchas learned

- The app is operated through the browser; code is built and shipped via the
  GitHub tools. Walk through any Vercel/Entra steps in detail.
- Verify changes build (`npm install` + `vite build`) before pushing.
- The MSAL CDN is dead (deprecated at v3); the npm package via Vite is the only
  supported path — no `<script>`-tag CDN approach.
- Redirect URIs must match exactly (no trailing slash); **SPA** platform, not
  Web. Launch the app from the production domain, not a hashed deployment URL.
- **Safety default:** prefer reversible operations (move, label, archive) over
  permanent deletion, and confirm before anything destructive. This tool only
  ever *moves* mail into Delete Review — to undo, move it back in Outlook.

## Tech

Vite · `@azure/msal-browser` · Microsoft Graph v1.0 (`/me/messages`,
`/me/mailFolders`, `/$batch`). No backend.
