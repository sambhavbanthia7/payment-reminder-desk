# Shared tracker in OneDrive

The tracker Excel file lives in one OneDrive. The app opens it automatically on every
computer, saves each change back about 2 seconds later, and checks for other people's
changes every 30 seconds (and whenever the app window is focused).

The app only has **app-folder** access (`Files.ReadWrite.AppFolder`). It can see and change
files in `Apps › JAL Receivables Desk` in that OneDrive, and nothing else.

For now this uses a personal Microsoft account. Use a **new account made for this**
(e.g. a new outlook.com address), not someone's own. The tracker holds customer names,
emails and balances, and the account must be handed over if that person leaves.

## 1. Register the app (one time, about 10 minutes)

1. Sign in at https://entra.microsoft.com (or https://portal.azure.com) with that Microsoft account.
   If Microsoft says the account has no directory, it may ask you to create a free Azure account
   first. That can ask for phone and card verification. Registering an app is free.
2. Go to **App registrations → New registration**:
   - Name: `JAL Receivables Desk`
   - Supported account types: **Personal Microsoft accounts only**
   - Redirect URI: **Web**, `http://localhost:5173/api/onedrive/callback`
3. Copy the **Application (client) ID** from the Overview page.
4. Go to **Certificates & secrets → New client secret** and copy the **Value**, which is shown only once.
   Note its expiry date.
5. Go to **API permissions → Add a permission → Microsoft Graph → Delegated**, and add
   `Files.ReadWrite.AppFolder`, `User.Read` and `offline_access`.

## 2. Settings (`.dev.vars`)

```
ONEDRIVE_CLIENT_ID=<Application (client) ID>
ONEDRIVE_CLIENT_SECRET=<client secret value>
ONEDRIVE_TENANT=consumers
```

Remove the practice lines (`ONEDRIVE_LOGIN_URL`, `ONEDRIVE_GRAPH_URL`) and restart the dev server.

## 3. Connect and share the tracker

1. In the app, click **Connect OneDrive** and sign in with that account. This is done once for
   everyone, and the connection is stored on the app's server.
2. Open the tracker from this computer as usual, then choose **OneDrive → Put the open workbook
   into OneDrive**.
3. From now on every computer opens that file automatically.
4. To let colleagues open the file directly in Excel as well, share the `Apps/JAL Receivables Desk`
   folder from OneDrive on the web. The app picks up their edits.

## How changes from different places are handled

| Situation | What happens |
|---|---|
| A change in the app | Saved to OneDrive about 2 seconds later. The panel shows "Saved to OneDrive · time". |
| Someone changed the file elsewhere, and this computer has no unsaved changes | This computer loads the new version and says who changed it |
| Someone changed the file elsewhere, and this computer has unsaved changes | Nothing is overwritten. The panel offers **Open OneDrive version**, **Keep mine and replace** or **Download my copy**. |
| The file is locked (open for editing elsewhere) | The panel says so, and saving retries every 15 seconds |

Two people changing things at the same moment cannot be merged. The rule is still one
editor at a time, but a clash is now caught instead of silently losing work.

## Practice OneDrive (no Microsoft account needed)

```bash
npm run onedrive:practice
```

This runs a fake Microsoft sign-in and OneDrive at http://127.0.0.1:8787, with files in
`.onedrive-practice/`. Point `.dev.vars` at it with `ONEDRIVE_CLIENT_ID=practice-client`,
`ONEDRIVE_CLIENT_SECRET=practice-secret`, `ONEDRIVE_LOGIN_URL=http://127.0.0.1:8787` and
`ONEDRIVE_GRAPH_URL=http://127.0.0.1:8787/v1.0`. The override URLs are only accepted for localhost.

- To imitate a colleague's edit, change the `.xlsx` in that folder.
- To imitate a locked file, create `<file name>.lock` next to it.

## Later: moving to JAL's own Microsoft 365

Register the app in JAL's tenant and set `ONEDRIVE_TENANT` to the tenant ID. Then connect
the work account and put the tracker into its OneDrive the same way.

At deployment, add `https://<site>/api/onedrive/callback` as a redirect URI and set the values
as Cloudflare secrets. All OneDrive routes require the app's sign-in.
