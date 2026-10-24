# JAL Receivables Desk

A web app for JAL's accounts team to follow up on outstanding invoices from the Excel debtor
tracker: review ageing and risk, record collections, send payment reminders and balance
confirmations from the finance mailbox, and keep the full history inside the Excel file.

## What it does

- **Sign-in with the finance mailbox.** Staff sign in with the mailbox address and password used in
  Outlook. The password is checked with the mail server and stored encrypted on the app's server.
  Only the addresses in `EMAIL_ALLOWED_ACCOUNTS` can sign in.
- **Works directly on the Excel tracker.** In Chrome or Edge, **Open debtor tracker** opens the
  `.xlsx` file and every change is saved straight back into it about a second later. The app
  remembers the file between visits, picks up edits made in Excel, and asks before overwriting
  anything changed elsewhere. Other browsers can download an updated copy instead.
- **Collections and records.** Record or correct payments (Receipts sheet), update customer emails,
  flag, hold or escalate invoices, and add notes and sales remarks.
- **Analysis.** Overview, ageing (overdue and invoice age), risk and escalation, data checks, search
  and filters, and a review date.
- **Payment reminders.** An automatic queue (due within 5 days or overdue, not on hold, not reminded
  in the last 7 days), with previews and frozen draft snapshots. Drafts are sent over SMTP from the
  signed-in mailbox. In test mode every email goes only to the test inboxes.
- **Safeguards.** A draft is blocked if its balance or contact changed, the same draft is never sent
  twice, sending stays under the host's hourly limit, and a hidden copy goes to the finance inbox.
- **Column Q.** Live reminders append `dd.mm.yy MAIL` to *Last Reminder Sent*, keeping its earlier
  text.
- **Balance confirmations.** Whole-balance requests are sent like reminders, and the customer's
  Confirmed or Disputed answer is recorded.
- **History in the workbook.** The Activity Log, Reminder Log, Followup State, Balance Confirmations,
  Desk Settings and Desk Summary sheets are added to the Excel file.
- **Optional: a shared tracker in OneDrive.** This is built but switched off. See
  `docs/ONEDRIVE_SETUP.md`.

## Run locally

Requires Node.js 22.13 or newer.

```
npm ci
cp .dev.vars.example .dev.vars    # then fill in the values
npm run dev                       # http://localhost:5173
```

To test without the real mailbox, run `npm run mail:catcher` and point `.dev.vars` at it, as
described in `docs/EMAIL_SETUP.md`.

## Commands

| Command | Purpose |
|---|---|
| `npm run dev` | Development server |
| `npm test` | All automated tests |
| `npm run lint` | ESLint |
| `npm run build` | Production build |
| `D1_DATABASE_ID=… D1_DATABASE_NAME=… npm run deploy` | Build and deploy to Cloudflare (`docs/DEPLOY.md`) |
| `npm run mail:catcher` | Practice mail server (SMTP + IMAP) on this computer |
| `npm run onedrive:practice` | Practice OneDrive on this computer |

## How it is built

- React 19 with Vinext (a Next.js-compatible framework on Vite), running on Cloudflare Workers.
- **Cloudflare D1** stores sign-in sessions, send records (to prevent duplicates and count the hourly
  limit) and the optional OneDrive connection. The receivables data itself lives only in the Excel
  file and in the browser's recovery copy.
- Email goes over SMTP with TLS, using a small built-in client (`lib/mail/`) on Cloudflare TCP sockets.
- The workbook is read and written in the browser (`lib/debtors-workbook.ts`). The original sheets,
  formulas and formatting are preserved, and only the input cells, formula results and log sheets are
  updated.

| Path | Contents |
|---|---|
| `app/debtors-dashboard.tsx` | The dashboard: pages, forms, sending, file saving |
| `app/login-page.tsx`, `app/api/auth/*` | Sign-in and sign-out |
| `app/api/email/*` | Email status and sending |
| `app/api/onedrive/*` | Optional OneDrive sync |
| `lib/debtors.ts` | Balances, ageing, stages, risk, queue, messages |
| `lib/debtors-workbook.ts` | Excel import and export |
| `lib/email-dispatch.ts` | Test-mode routing, stale-draft check, recording results |
| `lib/excel-file.ts`, `lib/debtors-storage.ts` | Saving into the file, and the browser recovery copy |
| `lib/mail/` | MIME, SMTP and IMAP clients |
| `lib/server/` | Settings, sessions, encryption, database, OneDrive (server only) |
| `scripts/` | Practice servers and the deploy script |
| `docs/` | Email, OneDrive and deployment guides |

## Limits

- Use Chrome or Edge to save into the Excel file. Close the file in Excel while the app is saving.
- One person should edit the tracker at a time. Clashes are detected, not merged.
- "Accepted by the mail server" means the server took the email. Bounces arrive in the mailbox.
