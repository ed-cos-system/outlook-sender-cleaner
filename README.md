# Sender Sweep — Outlook Sender Cleaner

A tiny single-page web app: type in a sender's email address and it moves
**every message from that sender** out of your Outlook mailbox into a folder
called **Delete Review**. Nothing is ever permanently deleted — Delete Review
is just a holding pen you can scan before emptying it yourself.

Built for personal Microsoft accounts (`@msn.com`, `@outlook.com`,
`@hotmail.com`, `@live.com`).

## How it works (and why there are no secrets)

- It's a static SPA (Vite). It signs you in with **MSAL** using the OAuth 2.0
  **Authorization Code flow with PKCE** — the standard for browser apps.
- Because it's a public client with PKCE, **there is no client secret** to
  store. The only config is your app's **client id**, which is not sensitive.
- All mail access happens in your browser with your own delegated token via
  the Microsoft Graph API. No server, no database, nothing stored anywhere
  except the MSAL token cache in your browser.

## One-time setup

### 1. Register an app to get a client id

This is the only part that has to be done by hand — it can't be scripted for
you, because it creates the identity your sign-in trusts.

1. Go to <https://entra.microsoft.com> (or the App registrations blade in
   <https://portal.azure.com>) and sign in with your personal Microsoft
   account.
2. **App registrations → New registration.**
3. Name: `Outlook Sender Cleaner` (anything is fine).
4. **Supported account types:** choose **Personal Microsoft accounts only**.
   (This must match the app's `authority`, which is set to `/consumers`.)
5. **Redirect URI:** platform = **Single-page application (SPA)**, and enter
   your production URL with **no trailing slash**, e.g.
   `https://outlook-sender-cleaner.vercel.app`
6. Click **Register**, then copy the **Application (client) ID**.
7. Open **Authentication** and add a second SPA redirect URI for local dev:
   `http://localhost:5173`
8. Open **API permissions → Add a permission → Microsoft Graph → Delegated
   permissions**, add **`Mail.ReadWrite`** and **`User.Read`**. (For personal
   accounts these are consented when you first sign in — no admin step.)

No client secret is needed. Leave that section alone.

### 2. Give the app your client id

Either option works (the id is public, so it can live in client code):

- **Recommended:** in Vercel → your project → **Settings → Environment
  Variables**, add `VITE_MS_CLIENT_ID` = your client id, then redeploy.
  (Vite inlines env vars at build time, so a redeploy is required after
  changing it.)
- **Or:** edit `src/config.js` and replace `PASTE_YOUR_CLIENT_ID_HERE`.

### 3. Deploy

This repo is already connected to Vercel. Vercel auto-detects Vite
(`build` → `vite build`, output `dist`). Push to the default branch and it
deploys. Make sure your live Vercel URL exactly matches the SPA redirect URI
you registered in step 1.

## Local development

```bash
npm install
VITE_MS_CLIENT_ID=your-client-id npm run dev   # http://localhost:5173
npm run build                                  # production build into dist/
```

## Using it

1. **Sign in with Microsoft** — pick the account that receives the mail.
2. Type a **sender address** (e.g. `noreply@example.com`).
3. Choose **Inbox only** (default) or **All folders**.
4. **Find messages** — it shows the count and a preview of subjects.
5. **Move to Delete Review** — it creates the folder if needed and moves
   everything in batches.

## About your two accounts

`f14stalker@msn.com` and `stalkersignup@outlook.com` both land in the same
inbox. Two cases:

- **They're aliases on one Microsoft account.** Sign in once and you're done —
  all mail lives in that single mailbox, and the tool filters by the *sender*
  regardless of which alias received it.
- **They're two separate accounts.** Use **Sign out**, then sign in to the
  other account and run the sweep again. (Sign-in uses "select account" so you
  can switch.)

If you're not sure which case you're in, just try one sign-in first; if some
mail is missing, it's the second case.

## Undo / safety

The tool only ever **moves** messages into **Delete Review** — it never
deletes. To undo, open Outlook and move them back, or delete the folder's
contents yourself once you've reviewed them.

## Tech

Vite · `@azure/msal-browser` · Microsoft Graph v1.0 (`/me/messages`,
`/me/mailFolders`, `/$batch`). No backend.
