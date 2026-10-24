# Email sending setup (cPanel mailbox over SMTP)

If Outlook on the finance computer uses POP (incoming port 110), copies saved to the server's Sent
folder never appear in Outlook. In that case leave `EMAIL_IMAP_HOST` blank and set `EMAIL_BCC` to the
finance mailbox instead, so each live email also lands in its inbox.

JAL's email is hosted with their website host (cPanel). Outlook is only the program staff
read it in. The Receivables Desk sends reminders the same way Outlook does: it connects to
the host's outgoing mail server (SMTP) with the mailbox address and password. It then saves
a copy in the mailbox's Sent folder over IMAP, so the team still sees every reminder in Outlook.

The request sent to IT is `JAL Receivables Desk - Email details for IT.docx` in the workspace root.

## Settings (`.dev.vars` locally, Cloudflare secrets when deployed)

Template: `.dev.vars.example`. Never commit `.dev.vars`.

| Setting | Meaning |
|---|---|
| `EMAIL_ALLOWED_ACCOUNTS` | Mailbox address(es) allowed to sign in. Reminders are sent from the signed-in account. |
| `SESSION_SECRET` | 32+ random characters that encrypt signed-in passwords on the server (`openssl rand -base64 48`) |
| `EMAIL_SMTP_HOST`, `EMAIL_SMTP_PORT` | Outgoing server. 465 means SSL and 587 means STARTTLS (worked out from the port). 25 is refused. |
| `EMAIL_SMTP_SECURITY` | Optional override: `ssl` or `starttls` |
| `EMAIL_IMAP_HOST`, `EMAIL_IMAP_PORT` | Incoming server, used only to save copies in Sent. Blank host means no copies. |
| `EMAIL_BCC` | Hidden copy of every **live** email (e.g. the finance mailbox), so the team sees it in Outlook whether Outlook uses POP or IMAP. Test sends skip it. |
| `EMAIL_SENT_FOLDER` | Optional. The Sent folder is found automatically (`\Sent` flag, then `INBOX.Sent`, `Sent`, …). |
| `EMAIL_FROM_NAME`, `EMAIL_REPLY_TO` | Display name customers see; optional different reply address |
| `EMAIL_HOURLY_LIMIT` | The host's hourly limit. The app refuses to go over it (default 100). |
| `EMAIL_MODE`, `EMAIL_TEST_RECIPIENT` | `test` (default) sends everything to the test inbox. `live` emails customers. |

## Signing in

The app opens on a sign-in page. Staff enter the mailbox address and password, the same ones used
in Outlook. The app checks them by logging in to the SMTP server; nothing is sent during the check.
On success it keeps the password **encrypted** (AES-GCM, `SESSION_SECRET`) in D1 `app_sessions` for
30 days and gives the browser only a random HttpOnly session cookie. The password is never put in
`.dev.vars`, the browser or the Excel file.

- Only `EMAIL_ALLOWED_ACCOUNTS` can sign in.
- After 5 wrong passwords in 15 minutes, all sign-ins pause, so the host's brute-force protection
  does not block the mailbox.
- If the mail server later rejects the saved password (IT changed it), everyone on that account is
  signed out and asked to sign in again.
- Everything else (email and OneDrive routes) requires a signed-in session. This is also the site's
  lock for deployment.

## Practice mail server (test without the real mailbox)

```bash
npm run mail:catcher
```

This runs a fake SMTP server on 127.0.0.1:2525 and a fake IMAP server on 127.0.0.1:1143. Both
accept any password and keep every email in `.mail-catcher/` (`outbox/` holds what was sent;
`sent-folder/` holds the copies). Point `.dev.vars` at it with `EMAIL_SMTP_HOST=127.0.0.1`,
`EMAIL_SMTP_PORT=2525`, `EMAIL_SMTP_SECURITY=none`, `EMAIL_IMAP_HOST=127.0.0.1`,
`EMAIL_IMAP_PORT=1143`, `EMAIL_IMAP_SECURITY=none`, and allow a practice address in `EMAIL_ALLOWED_ACCOUNTS`.
The practice server accepts any password, unless it was started with `MAIL_CATCHER_PASSWORD=…`. Security `none` is only allowed for
localhost. Restart the dev server after changing `.dev.vars`.

## Using the real mailbox

1. Fill in `.dev.vars` with IT's server values, keeping `EMAIL_MODE=test` and your own `EMAIL_TEST_RECIPIENT`.
2. Restart the dev server and sign in with the mailbox address and password. The header shows "Test mode → <your inbox>".
3. Prepare a draft and click **Send test**. Check your inbox, and the Sent folder of accounts@ in Outlook.
4. Settings errors (wrong password, wrong port) show in a toast in plain language.

## Going live

Set `EMAIL_MODE=live` and restart, only when the team is ready.

- Drafts are rechecked against current balances and contacts, and are never sent twice (`draft_sends` in D1).
- Live results mark drafts Accepted or Failed, count as the last reminder, and fill the Reminder Log
  columns Sent at, Sent by, Sent to and Send result.
- "Accepted by the mail server" does not prove delivery. Bounces arrive in the mailbox.
- If the mailbox password changes, update `EMAIL_PASSWORD` (or the Cloudflare secret).
