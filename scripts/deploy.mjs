// Build and deploy to Cloudflare. Needs `npx wrangler login` once, and the D1 database id in
// D1_DATABASE_ID (from `npx wrangler d1 create jal-receivables`). See docs/DEPLOY.md.
import { execSync } from "node:child_process"

if (!process.env.D1_DATABASE_ID) {
  console.error("Set D1_DATABASE_ID first, e.g. D1_DATABASE_ID=xxxxxxxx-... npm run deploy (see docs/DEPLOY.md).")
  process.exit(1)
}
process.env.D1_DATABASE_NAME ??= "jal-receivables"
execSync("npx vinext build", { stdio: "inherit" })
execSync("npx wrangler deploy", { stdio: "inherit" })
