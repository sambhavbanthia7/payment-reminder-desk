import vinext from "vinext";
import { defineConfig } from "vite";

// Local development uses this placeholder D1 id (Miniflare keeps local data under it).
// Deploys set D1_DATABASE_ID / D1_DATABASE_NAME to the real database (see docs/DEPLOY.md).
const LOCAL_DATABASE_ID = "00000000-0000-4000-8000-000000000000";

const workerConfig = {
  main: "./worker/index.ts",
  compatibility_flags: ["nodejs_compat"],
  d1_databases: [
    {
      binding: "DB",
      database_name: process.env.D1_DATABASE_NAME ?? "site-creator-d1",
      database_id: process.env.D1_DATABASE_ID ?? LOCAL_DATABASE_ID,
    },
  ],
};

export default defineConfig(async () => {
  // Keep Wrangler and Miniflare state inside the project.
  process.env.WRANGLER_WRITE_LOGS ??= "false";
  process.env.WRANGLER_LOG_PATH ??= ".wrangler/logs";
  process.env.MINIFLARE_REGISTRY_PATH ??= ".wrangler/registry";

  // Wrangler snapshots its log path while the Cloudflare plugin is imported.
  const { cloudflare } = await import("@cloudflare/vite-plugin");

  return {
    plugins: [
      vinext(),
      cloudflare({
        viteEnvironment: { name: "rsc", childEnvironments: ["ssr"] },
        inspectorPort: false,
        config: workerConfig,
      }),
    ],
  };
});
