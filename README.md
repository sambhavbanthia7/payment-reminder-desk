# Payment Reminder Desk

An internal accounts-receivable application for reviewing outstanding invoices,
resolving customer-data problems, manually selecting recipients, and sending
payment reminders through the signed-in employee's Microsoft Outlook mailbox.

> **Important:** The application is functionally complete for controlled testing,
> but deploying it is not the same as making it production-ready. Before using
> real customer data or sending real emails, complete every item in
> [Production-readiness requirements](#production-readiness-requirements).

## Contents

- [What the application does](#what-the-application-does)
- [Business logic](#business-logic)
- [Architecture](#architecture)
- [Current readiness](#current-readiness)
- [Prerequisites](#prerequisites)
- [Local development](#local-development)
- [Microsoft Outlook setup](#microsoft-outlook-setup)
- [Database setup](#database-setup)
- [Deployment](#deployment)
- [Production-readiness requirements](#production-readiness-requirements)
- [Pre-launch testing](#pre-launch-testing)
- [Operations and maintenance](#operations-and-maintenance)
- [Troubleshooting](#troubleshooting)
- [Project structure](#project-structure)

## What the application does

The employee workflow is:

1. Sign in with the Outlook account that will send reminders.
2. Configure the campaign:
   - first reminder window, default 7 days;
   - priority reminder window, default 3 days;
   - placeholder credit period, default 30 days;
   - email subject and body.
3. Upload a dues workbook and a customer-master workbook.
4. Review or manually change the detected column mappings.
5. Review:
   - invoice value, invoice count, and unique customer count in every summary;
   - invoices due within the first reminder window;
   - invoices due within the priority window;
   - invoices that are already overdue;
   - the exact source rows behind any summary by clicking its card;
   - customers that cannot be emailed automatically.
6. Enter a campaign-only replacement email for missing, invalid, or unmatched
   contacts when the correct address is known.
7. Manually choose the customers who should receive an email.
8. Click the send button.
9. Review successful sends, failures, and sample previews in the activity log.

The two supplied test workbooks are represented in the application as preloaded
sample data. Sample mode never sends real email.

## Business logic

### Due-date calculation

```text
Due date       = Invoice date + Credit days
Days remaining = Due date - Report date
```

If `Credit Days` is blank, the configured placeholder is used. The default
placeholder is 30 days. The interface flags every customer for whom the
placeholder was required.

Negative or invalid credit-day values are not replaced automatically; those
records are blocked and shown for manual review.

### Invoice eligibility

A row is treated as an open invoice only when:

- the trimmed transaction type equals `GST INVOICE`;
- the cleaned balance is greater than zero;
- the invoice date is parseable; and
- the row is not a subtotal, total, receipt, return, journal, or opening balance.

The XLSX parser handles:

- amounts stored as numbers;
- Indian-formatted amounts such as `1,29,715.00`;
- HTML-wrapped values;
- `DD/MM/YY` and `DD/MM/YYYY` dates;
- leading or trailing whitespace in source fields.

### Customer matching

```text
Dues workbook:   A/C Code
Customer master: Party Code
```

The normalized code is the join key. Customer names are not used for matching
because the dues export may append a location such as `>> Manali`.

The dues workbook is authoritative for invoice credit terms. The master is used
for customer identity and email address.

A valid email entered on the **Needs attention** page overrides the missing,
invalid, or unmatched contact for the current in-browser campaign and makes that
customer eligible for selection. It does not rewrite the uploaded master file,
and overrides are cleared when the master file is replaced or the test data is
restored.

### Reminder windows

- **Due within first window:** `0 <= Days Remaining <= First Threshold`
- **Due within priority window:** `0 <= Days Remaining <= Priority Threshold`
- **Overdue:** `Days Remaining < 0`

The application does not send automatically on a schedule. It prepares the
appropriate customer list, but an employee must select recipients and press the
send button.

### Email grouping

Invoices are grouped by party code. A customer selected in a particular view
receives one consolidated email containing the invoices currently included in
that view.

Supported email-template variables:

| Variable | Meaning |
| --- | --- |
| `{{customer_name}}` | Customer name |
| `{{party_code}}` | Account/party code |
| `{{invoice_count}}` | Number of invoices in the email |
| `{{total_due}}` | Consolidated outstanding amount |
| `{{invoice_list}}` | Invoice numbers, balances, and due dates |
| `{{reminder_stage}}` | Selected reminder window |

## Architecture

| Layer | Implementation |
| --- | --- |
| UI | React 19, Next-compatible App Router, Tailwind CSS |
| Runtime/build | Vinext and Vite |
| Production runtime | Cloudflare Worker-compatible ESM |
| Spreadsheet processing | Browser-side XLSX ZIP/XML parser |
| Authentication | Microsoft authorization-code flow with PKCE |
| Email | Microsoft Graph `/me/sendMail` |
| Persistence | Cloudflare D1 |
| Session cookie | Random server-side session ID; `Secure`, `HttpOnly`, `SameSite=Lax` |

Workbook contents are processed in the employee's browser. The current
implementation does not upload or persist the workbook files themselves.
Microsoft access and refresh tokens are stored server-side in D1.

## Current readiness

### Implemented

- Responsive dashboard and employee workflow
- Preloaded sample campaign
- Real `.xlsx` upload
- Automatic sheet and header detection
- Manual column mapping
- Configurable reminder windows and credit-days placeholder
- Due, upcoming, priority, and overdue calculations
- Invoice and unique-customer counts on every KPI card
- Click-through source-row previews for every KPI bucket
- Missing, invalid, and unmatched customer flags
- Campaign-only manual email overrides for contact-data exceptions
- Manual recipient selection
- Microsoft OAuth with PKCE
- Microsoft Graph email sending
- Token refresh
- D1-backed session and send-log storage
- Sample previews that cannot send email
- Production build validation

### Not yet sufficient for unrestricted production use

- Organization-level authorization/role enforcement
- App-level encryption for Microsoft tokens stored in D1
- Database migrations managed outside request handling
- Duplicate-send prevention/idempotency
- Rate limiting and send-volume controls
- CSRF/origin enforcement on write endpoints beyond SameSite cookies
- File-size, row-count, and decompression limits
- Centralized monitoring and alerting
- Formal backup, retention, and incident-response policies
- Automated tests for the business calculation layer
- User-acceptance testing against representative production exports

These are listed as required launch work below.

## Prerequisites

### Accounts and services

- A Cloudflare account capable of running Workers and D1, or a ChatGPT Sites
  environment that provides the equivalent Worker and D1 bindings
- A Microsoft 365 work or school mailbox with permission to send email
- Access to Microsoft Entra ID to create an app registration
- A production HTTPS hostname
- Access to the organization's DNS configuration
- A designated finance/application owner

### Development tools

- Node.js 22.13 or newer
- npm
- Linux or WSL with `bash`, `flock`, `curl`, `sha256sum`, and GNU
  `timeout`
- Git
- Wrangler CLI when deploying directly to Cloudflare

The supplied install and build scripts target Linux/WSL. On macOS or Windows
without WSL, either run the project in a Linux container or adapt the wrapper
scripts.

## Local development

### 1. Install

```bash
git clone <your-repository-url>
cd payment-reminder-desk
cp .env.example .env.local
npm ci
```

### 2. Set local environment values

```dotenv
MICROSOFT_CLIENT_ID=
MICROSOFT_CLIENT_SECRET=
MICROSOFT_TENANT_ID=organizations
```

The dashboard and sample workspace can be used without Microsoft credentials.
Outlook sign-in and real email sending remain disabled until the client ID is
present.

### 3. Run

```bash
npm run dev
```

### 4. Quality checks

```bash
npm run lint
npm test
```

`npm test` runs the production build before the included structural tests.

### Local OAuth warning

Authentication cookies are always marked `Secure`. Standard HTTP localhost
will therefore not retain the Outlook session in most browsers. Use an HTTPS
development hostname/tunnel whose callback URI is registered in Entra, or add a
carefully environment-gated development-cookie option. Never weaken the cookie
in production.

## Microsoft Outlook setup

The code uses delegated Microsoft Graph permissions and sends through:

```text
POST https://graph.microsoft.com/v1.0/me/sendMail
```

This means emails are sent from the mailbox of the employee who signs in.

### 1. Create the Entra app

In Microsoft Entra admin center:

1. Open **App registrations**.
2. Select **New registration**.
3. Use a clear internal name, such as `Payment Reminder Desk - Production`.
4. For a company-only deployment, select accounts in the organization's
   directory only.
5. Register the application.
6. Copy:
   - **Application (client) ID**;
   - **Directory (tenant) ID**.

For production, use the actual tenant ID rather than `organizations`. This
prevents employees from unrelated Microsoft tenants from authenticating.

### 2. Add redirect URIs

Under **Authentication**, add a Web redirect URI for every environment:

```text
https://YOUR_PRODUCTION_DOMAIN/api/auth/microsoft/callback
https://YOUR_STAGING_DOMAIN/api/auth/microsoft/callback
```

The URI must match the deployed scheme, hostname, and path exactly. Do not use
a wildcard.

### 3. Create a credential

Under **Certificates & secrets**:

1. Create a client secret for initial deployment, or adopt your organization's
   approved certificate-based credential process.
2. Store the secret immediately in the deployment secret manager.
3. Record the owner and expiry date.
4. Create a rotation reminder before expiry.

Never commit the secret to Git, `.env.example`, Wrangler configuration, CI
logs, screenshots, or tickets.

### 4. Configure permissions

Under **API permissions**, add delegated Microsoft Graph permissions:

- `User.Read`
- `Mail.Send`

The application also requests `openid`, `profile`, and `offline_access`
during the OAuth flow. Depending on tenant policy, an Entra administrator may
need to grant consent.

Do not add application-wide `Mail.Send` unless the business has explicitly
decided to use unattended sending and has completed the corresponding security
review. This application is designed for delegated, employee-approved sending.

### Shared mailbox limitation

The current route sends from `/me`. It does not send from a shared mailbox.
To support a shared accounts-receivable mailbox, implement and test:

- the appropriate Microsoft Graph endpoint for the shared mailbox;
- delegated `Mail.Send.Shared` where applicable; and
- Exchange **Send As** or **Send on Behalf** permission for the signed-in user.

Do not assume that connecting a user automatically grants access to a shared
mailbox.

### Environment variables

| Variable | Required | Secret | Description |
| --- | --- | --- | --- |
| `MICROSOFT_CLIENT_ID` | Yes | Treat as configuration | Entra application/client ID |
| `MICROSOFT_CLIENT_SECRET` | Yes for the current confidential-web-app setup | Yes | Entra client secret |
| `MICROSOFT_TENANT_ID` | Recommended | No | Tenant UUID; defaults to `organizations` |

The Worker runs with Node.js compatibility, so these values are read through
`process.env`.

## Database setup

The application expects a Cloudflare D1 binding named exactly:

```text
DB
```

### Current tables

#### `outlook_sessions`

Stores:

- random session ID;
- Microsoft access token;
- Microsoft refresh token;
- token expiry;
- mailbox email and display name;
- creation and update timestamps.

#### `send_logs`

Stores:

- campaign ID;
- party code and customer name;
- recipient email;
- reminder stage;
- invoice numbers and count;
- consolidated amount;
- send status and error;
- sending mailbox;
- timestamp.

### Current initialization behavior

`lib/server/database.ts` runs `CREATE TABLE IF NOT EXISTS` on demand. This is
convenient for development and first deployment, but schema changes should not
be managed inside normal request processing in production.

Before launch:

1. Create versioned SQL migrations.
2. Apply migrations in CI/CD before promoting application code.
3. Use separate databases for development, staging, and production.
4. Test rollback and restore procedures.
5. Define retention for session and send-log records.
6. Delete expired sessions with a scheduled maintenance job.

### Token protection

The session table contains Microsoft bearer and refresh tokens. Provider-level
encryption at rest is not a complete application security boundary. Before
production, encrypt token values at the application layer using a managed key,
restrict database access, redact tokens from logs, and document the key-rotation
procedure.

## Deployment

The repository is currently shaped for a Cloudflare Worker runtime with D1.
Deploying to Vercel, a traditional Node server, AWS, or another platform
requires replacing the Cloudflare environment and D1 adapter.

### Option A: ChatGPT Sites

The included `.openai/hosting.json` declares:

```json
{
  "d1": "DB",
  "project_id": "<managed by Sites>",
  "r2": null
}
```

For a new Site:

1. Create/import the Site.
2. Ensure the logical D1 binding is `DB`.
3. Add the Microsoft variables through the Site's environment-variable
   controls.
4. Keep the Site private or restricted to the finance team.
5. Build and publish a checkpoint.
6. Register the final Site hostname as the Entra redirect URI.
7. Complete the production checklist before enabling real sends.

Do not reuse the project ID from another deployment.

### Option B: Cloudflare Workers and D1

#### 1. Authenticate and create D1

```bash
npx wrangler login
npx wrangler d1 create payment-reminder-production
```

Save the returned database ID.

#### 2. Add a Wrangler configuration

The Sites checkout intentionally does not contain a reusable
`wrangler.jsonc`. For an independent Cloudflare deployment, create one based
on the emitted Vinext worker and your real D1 resource:

```jsonc
{
  "$schema": "./node_modules/wrangler/config-schema.json",
  "name": "payment-reminder-desk",
  "main": "./dist/server/index.js",
  "compatibility_date": "2026-08-29",
  "compatibility_flags": ["nodejs_compat"],
  "assets": {
    "directory": "./dist/client",
    "binding": "ASSETS",
    "run_worker_first": true
  },
  "d1_databases": [
    {
      "binding": "DB",
      "database_name": "payment-reminder-production",
      "database_id": "REPLACE_WITH_DATABASE_ID"
    }
  ],
  "observability": {
    "enabled": true
  }
}
```

Confirm the build's emitted server entry and static-assets directory before
deploying; keep the configuration aligned with the pinned Vinext version.

#### 3. Store secrets

```bash
npx wrangler secret put MICROSOFT_CLIENT_ID
npx wrangler secret put MICROSOFT_CLIENT_SECRET
npx wrangler secret put MICROSOFT_TENANT_ID
```

Cloudflare secrets must not be placed directly in `wrangler.jsonc`.

#### 4. Build and deploy

```bash
npm ci
npm run build
npx wrangler deploy
```

#### 5. Configure the production hostname

1. Attach the final custom domain.
2. Confirm HTTPS is active.
3. Add the final callback URI to the Entra app.
4. Sign in again after changing domains because cookies are host-specific.

### Other hosting platforms

To run elsewhere, replace at least:

- `cloudflare:workers` environment access;
- the D1 implementation in `lib/server/database.ts`;
- the Worker entry/build and static-asset wiring;
- secret injection; and
- platform-specific cookie/proxy configuration.

Preserve the APIs exported by the database adapter to minimize application
changes.

## Production-readiness requirements

Complete these before real customer emails are enabled.

### 1. Access control

Outlook authentication identifies a Microsoft user, but the application does
not currently verify that the user belongs to an approved finance group.

Required:

- use a single-tenant Entra app;
- place the application behind Cloudflare Access, Sites access controls, or an
  equivalent identity-aware proxy;
- enforce an application-side allowlist or Entra group/role claim on every
  write endpoint;
- deny unknown users before workbook data is displayed;
- define separate viewer and sender roles if needed.

### 2. Secret and token security

- Keep the Microsoft client secret in the platform secret store.
- Encrypt access and refresh tokens before writing them to D1.
- Rotate application credentials and encryption keys.
- Redact authorization headers, tokens, workbook values, and email bodies from
  logs.
- Set a maximum session lifetime and remove expired sessions.
- Revoke sessions when an employee leaves the team.

### 3. Prevent accidental duplicate sends

The current log records sends but does not block duplicates.

Add:

- a stable idempotency key such as campaign + party code + invoice numbers +
  reminder stage;
- a unique database constraint;
- a server-side transaction that reserves the key before calling Graph;
- a clear warning if some invoices were already emailed;
- an explicit confirmation screen showing recipients, invoice count, and total.

### 4. Validate and limit uploads

The XLSX parser runs in the browser. Add:

- maximum file size;
- maximum compressed and decompressed size;
- maximum sheet, row, and column count;
- required-column validation;
- duplicate party-code detection;
- duplicate invoice-number detection;
- report-date consistency checks;
- rejection of unsupported `.xls` files;
- clear handling for malformed ZIP/XML content;
- test cases based on sanitized production exports.

The application currently supports `.xlsx`, not legacy binary `.xls`.

### 5. Sending controls

- Add per-user and per-campaign rate limits.
- Cap recipients per batch.
- Decide whether partial success should stop or continue a batch.
- Implement retry only for safe, transient failures.
- Never automatically retry an ambiguous Graph response without checking the
  idempotency record.
- Record the Microsoft request/result correlation ID where available.
- Consider a two-person approval for large campaigns or high balances.

### 6. API protection

- Verify `Origin` on state-changing requests.
- Add an explicit CSRF token for the send action.
- Validate request payloads with a strict schema.
- Reject unknown fields and excessive body sizes.
- Apply security headers and a restrictive Content Security Policy.
- Keep `Secure`, `HttpOnly`, and `SameSite` cookie attributes.

### 7. Audit and compliance

Decide and document:

- who is allowed to view customer balances;
- who may send reminders;
- how long email logs are retained;
- whether email bodies must be retained;
- how corrections are recorded;
- how access requests and deletions are handled;
- applicable Indian privacy, contractual, and financial-control requirements;
- who approves wording and customer escalation policy.

Avoid putting sensitive financial data into general-purpose application logs.

### 8. Monitoring and alerting

Add:

- server error tracking;
- structured audit events;
- Graph failure-rate monitoring;
- authentication failure monitoring;
- D1 error and latency monitoring;
- alerts for unusual send volume;
- uptime checks for the dashboard and callback route;
- a user-visible request/reference ID for support.

### 9. Backups and recovery

- Configure D1 backup/export procedures.
- Test restore into a separate database.
- Document rollback of application versions.
- Keep database changes backward-compatible during deployment.
- Define recovery-point and recovery-time objectives.

### 10. Staging and separation of duties

Use separate:

- Entra app registrations;
- redirect URIs;
- client secrets;
- D1 databases;
- domains; and
- access policies

for development, staging, and production.

Never use a production mailbox while testing.

## Pre-launch testing

### Build checks

```bash
npm ci
npm run lint
npm test
```

### Workbook tests

- Correct dues and master files
- Swapped files
- Missing required columns
- Columns in a different order
- Headers after introductory rows
- Blank credit days
- Negative credit days
- Invalid dates
- Indian-formatted amounts
- HTML-wrapped amounts
- Zero and negative balances
- Duplicate party codes
- Duplicate invoice numbers
- Missing and invalid email addresses
- Customer absent from master
- Large but permitted workbook
- Malformed or password-protected workbook

### Calculation reconciliation

For at least one sanitized production export:

1. Recalculate totals independently.
2. Verify every due-date bucket.
3. Verify the 30-day placeholder cases.
4. Reconcile customer totals to the accounting source.
5. Have an accounts employee sign off the results.

### Outlook tests

- Successful sign-in
- Tenant restriction
- Consent behavior
- Sign-out
- Access-token expiry and refresh
- Client-secret expiry behavior
- Single send to an internal test mailbox
- Multiple customer sends
- Graph partial failure
- Sent Items entry
- Mailbox sending limits
- Revoked consent/session

### Security tests

- Anonymous access
- Unauthorized Microsoft user
- CSRF attempt
- Oversized JSON body
- Oversized workbook
- Malicious XLSX ZIP/XML payload
- HTML/script content in customer fields
- Secret leakage in logs
- Session reuse after logout
- Direct calls to send and log endpoints

## Operations and maintenance

### Daily

- Review failed sends.
- Review unmatched customers and invalid emails.
- Confirm unusual send volumes with finance.

### Monthly

- Review user access.
- Review send-log retention.
- Confirm backups are completing.
- Sample and reconcile reminder activity.

### Before credential expiry

- Create the replacement Entra secret.
- Update the deployment secret.
- Verify sign-in and token refresh.
- Revoke the old secret.

### When the workbook format changes

1. Obtain a sanitized sample.
2. Add or update header aliases in `lib/reminders.ts`.
3. Test parsing in `lib/xlsx.ts`.
4. Reconcile calculations.
5. Deploy to staging before production.

## Troubleshooting

### Outlook button says credentials are not configured

Check that `MICROSOFT_CLIENT_ID` is present in the deployed runtime. The
status endpoint intentionally reports Outlook as unavailable when it is absent.

### Entra reports a redirect URI mismatch

Register:

```text
https://THE_EXACT_CURRENT_HOST/api/auth/microsoft/callback
```

Check scheme, subdomain, path, trailing slash, and environment.

### Sign-in completes but the dashboard still shows disconnected

- Ensure the site is served over HTTPS.
- Ensure the D1 `DB` binding exists.
- Verify D1 writes are succeeding.
- Check that the browser accepts the secure session cookie.
- Confirm the app secret has not expired.

### Graph returns 401

- Reconnect Outlook.
- Confirm token refresh succeeds.
- Confirm the Entra secret is valid.
- Confirm the tenant ID matches the app registration.

### Graph returns 403

- Confirm delegated `Mail.Send` was granted.
- Check whether tenant admin consent is required.
- Confirm the signed-in user has a usable mailbox.
- For a shared mailbox, check the additional Graph and Exchange permissions.

### Workbook cannot be read

- Confirm the file is an unencrypted `.xlsx`.
- Confirm it contains a tabular sheet.
- Check file/ZIP integrity.
- Use the column-mapping screen if names differ.
- Legacy `.xls` files must be exported as `.xlsx`.

### Customer is blocked

The dashboard blocks automatic selection when:

- the party code is absent from the master;
- the email is blank;
- the email format is invalid; or
- credit days are negative/invalid.

Blank credit days are allowed using the configured placeholder and remain
flagged for review.

## Project structure

```text
app/
  api/
    auth/microsoft/     Outlook OAuth start, callback, status, logout
    outlook/send/       Microsoft Graph sending
    logs/               Activity-log retrieval
  payment-reminder-dashboard.tsx
lib/
  reminders.ts          Calculation, matching, grouping, templates
  sample-data.ts        Preloaded test fixture
  xlsx.ts               Browser-side XLSX parser
  server/
    database.ts         D1 sessions and activity logs
    microsoft.ts        OAuth, PKCE, cookies, token refresh
worker/
  index.ts              Cloudflare Worker entry point
components/ui/          Vendored UI primitives
.openai/hosting.json    Sites runtime bindings
vite.config.ts          Vinext/Vite/Cloudflare setup
```

## Official references

- [Microsoft authorization-code flow with PKCE](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-auth-code-flow)
- [Microsoft Graph sendMail](https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0)
- [Microsoft Graph permissions reference](https://learn.microsoft.com/en-us/graph/permissions-reference)
- [Sending from another user/shared mailbox](https://learn.microsoft.com/en-us/graph/outlook-send-mail-from-other-user)
- [Cloudflare D1 getting started](https://developers.cloudflare.com/d1/get-started/)
- [Cloudflare D1 local development](https://developers.cloudflare.com/d1/best-practices/local-development/)
- [Cloudflare Worker static assets](https://developers.cloudflare.com/workers/static-assets/)
- [Cloudflare Worker secrets](https://developers.cloudflare.com/workers/configuration/secrets/)

## License

No license has been selected. Add an approved `LICENSE` file before sharing
the repository outside the organization.
