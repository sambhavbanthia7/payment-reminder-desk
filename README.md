# JAL Receivables Desk

A browser-based debtor follow-up dashboard for the updated JAL tracker. Open one
`.xlsx` file containing `Debtors Data` and `Receipts`; no column mapping or
Microsoft sign-in is required for the current workflow.

## Run locally

```powershell
npm ci
npm run dev
```

Open http://localhost:5173 and choose **Open debtor tracker**. The tracker is
processed in the browser. The dashboard does not call the server database or
mail APIs. Existing Outlook code is retained for a later integration.

On Windows, the original Bash build wrapper requires WSL. A direct application
build and focused checks are available with:

```powershell
npx vinext build
npx eslint app/debtors-dashboard.tsx lib/debtors.ts lib/debtors-workbook.ts lib/debtors-storage.ts
node --test tests/debtors.test.mjs tests/log-export.test.mjs
```

## Dashboard

- Overview: outstanding, overdue exposure, critical customers, reminder queue,
  recent allocated collections, and top sales owners.
- Reminder queue: customers with invoices overdue or due within the configured
  pre-due window. Review contacts, select customers, preview and prepare drafts.
- Balance confirmation: whole open customer balances, confirmation drafts, and
  recorded customer responses or disputes.
- Invoice register: searchable invoice amounts, dates, receipts, outstanding,
  status, source row and an activity timeline. Record part payments, correct
  receipts, update customer contacts, add notes, flag, hold or escalate.
- Ageing: separate overdue and invoice-age analysis, upcoming due dates, and
  exposure by area.
- Risk: Critical / High / Watch / Standard customer exposure.
- History: activity audit, immutable message snapshots, confirmation responses,
  and original Last Reminder Sent entries.
- Data checks: missing contacts, duplicates, credit entries, overpayments, and
  unmatched receipts. Workbook summary limitations remain visible.

## Calculation rules

The fixed template is validated by sheet name and column order. Formula results
for balances and ageing are recomputed from invoice inputs and receipts, rather
than trusting the original Dashboard's cached totals or limited ranges.

Receipt allocation matches both party name and invoice number, case
insensitively, preserving whitespace as Excel SUMIFS does. Unmatched receipts
are shown separately. A receipt never silently allocates to another invoice.

Outstanding = max(0, invoice amount - allocated receipts). Negative credit rows
are shown separately and are not netted against unrelated invoices. Overpaid
invoices are flagged. Payment entry prevents a new allocation from exceeding an
invoice's available balance. Correcting a receipt retains the previous values
in the activity audit.

SOP stages follow the supplied formulas: less than 5 overdue days = Pre-Due /
Due; 5-14 = 1st Reminder; 15-29 = 2nd Reminder; 30-44 = 3rd Reminder / CFO;
45+ = Management Escalation. Risk is Critical at 45+ days or manual Management /
Legal escalation; High at 30-44 days, CFO escalation or manual flag; Watch at
15-29 days; otherwise Standard. Customer risk uses the highest open-invoice risk.

The initial pre-due window is 5 days and reminder repeat interval is 7 days.
Both can be changed in Email templates. Held invoices and drafts already
prepared for the same review date are excluded from the reminder queue. Imported
reminder dates suppress reminders within the configured repeat interval.

The review date controls ageing and queue eligibility. Balances include all
receipts, exactly as the source formulas do; changing the review date does not
produce a historical balance sheet. Confirmation requests always include the
customer's entire open balance, even when filters match only one invoice.

## Workbook-based history and saving

No backend database is needed for the dashboard. The working file and state are
autosaved to browser IndexedDB as a recovery draft. This is local to the browser
and origin, not a shared record. Clearing browser data removes that recovery copy.

**Save workbook** downloads an updated Excel copy. Reopen that copy next time to
continue with its complete history. Ordinary browser uploads cannot silently
overwrite the original file on disk.

**Save directly to a file** uses the File System Access API where supported
(typically desktop Chrome/Edge). Choose the intended working file once; later
saves update that selected file. External file changes are detected before a
subsequent write. Close the file in Excel before writing to it. Where this API
is unavailable, use the normal download action.

**Open debtor tracker** (Chrome or Edge) opens the Excel file directly, and every later change is
saved straight back into that same file about a second later. No second step, no "-updated" copy.
The app remembers the file. When it is reopened it reconnects automatically, or shows **Continue with
<file>** if the browser wants one click to allow saving again. If the file was edited in Excel in the
meantime, the app loads those edits. If the app also has unsaved changes, it asks which version to keep:
**Open the file's version**, **Keep the app's version** or **Download a copy**. A save blocked because
the file is open in Excel is retried every 15 seconds. Close the file in Excel while working in the app.
Browsers without direct file access fall back to uploading and **Download Excel + logs**.

Every workbook open creates an activity entry. The save panel distinguishes
browser-only changes from exported changes, shows the activity count and names
the log sheets. Exported files open on **Activity Log** so the history is visible.
The invoice list now exposes payment, contact and follow-up actions directly;
message previews can prepare a single draft without returning to row selection.

Keep one current working workbook and one editor at a time. This workflow does
not merge parallel edits made by multiple users or different browser tabs.

Exports preserve original workbook parts, formatting, formula definitions,
external links, historical reminder text, and source sheets. Contact and sales
remarks edits update the appropriate Debtors Data input cells. Payments update
or append Receipts records. Known calculated Debtors Data cells receive updated
cached results; Excel is instructed to recalculate formulas on opening.

The app adds or refreshes these sheets:

| Sheet | Purpose |
| --- | --- |
| Activity Log | Timestamped changes, including before and after values |
| Reminder Log | Recipient, invoices, balance snapshot, subject, body and draft status |
| Followup State | Manual flags, holds, escalations and notes keyed to invoice |
| Balance Confirmations | Customer responses, dates, balance snapshots and references |
| Desk Settings | Message templates and reminder cadence |
| Desk Summary | Full-data totals for the selected review date |

The supplied workbook's original Dashboard uses incomplete ranges and mismatched
summary categories. It is preserved; use the web dashboard or Desk Summary for
full-data totals. Original formula errors and external links are not silently
repaired. The app is not a general-purpose Excel calculation engine. Numeric
formula inputs need valid cached values, or supported simple arithmetic;
otherwise open and save the file in Excel before uploading it.

## Email sending

Prepared drafts are sent from JAL's cPanel mailbox over SMTP, and a copy is saved in its
Sent folder over IMAP. See `docs/EMAIL_SETUP.md` for the settings, and for the practice
mail server (`npm run mail:catcher`) that lets you test without the real mailbox.
Sending stays in test mode (every email goes to `EMAIL_TEST_RECIPIENT`) until
`EMAIL_MODE=live`. Drafts are rechecked before sending, are never sent twice, and stay
within the host's hourly limit. The app opens on a sign-in page: staff sign in with the mailbox address and password, and reminders are sent from that account.
Each live reminder also appends `dd.mm.yy MAIL` to the tracker's Last Reminder Sent column (Q), keeping its earlier text.

## Shared tracker in OneDrive

The tracker can live in OneDrive, so every computer opens the same file and changes save
automatically. Edits made elsewhere are picked up within about 30 seconds, and clashes are
caught rather than overwritten. See `docs/ONEDRIVE_SETUP.md`, including the practice OneDrive
(`npm run onedrive:practice`) for testing without a Microsoft account.

## Code

- `app/debtors-dashboard.tsx`: current dashboard, interactions and export controls.
- `app/debtors.css`: responsive dashboard styles.
- `lib/debtors.ts`: balances, grouping, SOP stages, queue, templates and edits.
- `lib/debtors-workbook.ts`: fixed-template import and preservation-based export.
- `lib/debtors-storage.ts`: browser recovery draft.
- `tests/debtors.test.mjs`: receipt, ageing, queue and contact calculation checks.

The earlier two-file dashboard remains in `app/payment-reminder-dashboard.tsx`
for reference and is no longer the home route.
