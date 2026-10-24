# Deploying to Cloudflare

Updating an existing deployment: rebuild with the same database settings it was first deployed with.

```
D1_DATABASE_ID=<database id> D1_DATABASE_NAME=<database name> npm run deploy
```

`npx wrangler d1 list` shows the database id and name.

## One time

1. Log in: `npx wrangler login`. This opens the browser; sign in to the Cloudflare account that will own the app.
2. Create the database: `npx wrangler d1 create jal-receivables`, then copy the `database_id` it prints.
3. Deploy: `D1_DATABASE_ID=<id> npm run deploy`. The site appears at
   `https://payment-reminder-desk.<your-subdomain>.workers.dev`.
4. Add the settings as secrets, using the same names as `.dev.vars.example`:
   `EMAIL_ALLOWED_ACCOUNTS`, `EMAIL_SMTP_HOST`, `EMAIL_SMTP_PORT`, `EMAIL_SMTP_SECURITY`, `EMAIL_BCC`,
   `EMAIL_FROM_NAME`, `EMAIL_HOURLY_LIMIT`, `EMAIL_MODE`, `EMAIL_TEST_RECIPIENT`, and a **new**
   `SESSION_SECRET` (`openssl rand -base64 48`). Put them in a temporary JSON file, run
   `npx wrangler secret bulk <file>`, then delete the file.
5. Open the site, sign in with the finance mailbox, and open the tracker in Chrome or Edge.

## Every later update

`D1_DATABASE_ID=<id> npm run deploy`

## Notes

- Keep `EMAIL_MODE=test` until the team has checked real test emails. Then set it to `live` with
  `npx wrangler secret put EMAIL_MODE`.
- Changing `SESSION_SECRET` signs everyone out.
- The mailbox password is never a secret here. Staff type it on the sign-in page.
- File auto-save needs Chrome or Edge, and works on the deployed https address as it does on localhost.
