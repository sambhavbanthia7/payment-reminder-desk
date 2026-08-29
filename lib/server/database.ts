import { env } from "cloudflare:workers"

type Statement = {
  bind: (...values: unknown[]) => Statement
  run: () => Promise<unknown>
  first: <T = Record<string, unknown>>() => Promise<T | null>
  all: <T = Record<string, unknown>>() => Promise<{ results?: T[] }>
}

type Database = {
  prepare: (query: string) => Statement
}

export type StoredSession = {
  id: string
  access_token: string
  refresh_token: string | null
  expires_at: number
  account_email: string
  account_name: string
}

export type StoredLog = {
  id: string
  party_code: string
  customer_name: string
  email: string
  reminder_stage: string
  invoice_numbers: string
  invoice_count: number
  total_due: number
  status: "sent" | "failed" | "previewed"
  error_message: string | null
  sent_at: string
}

export function getDatabase(): Database | null {
  return ((env as unknown as { DB?: Database }).DB ?? null)
}

export async function ensureDatabase(db: Database): Promise<void> {
  await db.prepare(`CREATE TABLE IF NOT EXISTS outlook_sessions (
    id TEXT PRIMARY KEY,
    access_token TEXT NOT NULL,
    refresh_token TEXT,
    expires_at INTEGER NOT NULL,
    account_email TEXT NOT NULL,
    account_name TEXT NOT NULL,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
  )`).run()
  await db.prepare(`CREATE TABLE IF NOT EXISTS send_logs (
    id TEXT PRIMARY KEY,
    campaign_id TEXT NOT NULL,
    party_code TEXT NOT NULL,
    customer_name TEXT NOT NULL,
    email TEXT NOT NULL,
    reminder_stage TEXT NOT NULL,
    invoice_numbers TEXT NOT NULL,
    invoice_count INTEGER NOT NULL,
    total_due REAL NOT NULL,
    status TEXT NOT NULL,
    error_message TEXT,
    sent_at TEXT NOT NULL,
    sent_by TEXT NOT NULL
  )`).run()
  await db.prepare("CREATE INDEX IF NOT EXISTS idx_send_logs_sent_at ON send_logs(sent_at DESC)").run()
}

export async function readSession(id: string): Promise<StoredSession | null> {
  const db = getDatabase()
  if (!db) return null
  await ensureDatabase(db)
  return db.prepare("SELECT id, access_token, refresh_token, expires_at, account_email, account_name FROM outlook_sessions WHERE id = ?")
    .bind(id).first<StoredSession>()
}

export async function writeSession(session: StoredSession): Promise<void> {
  const db = getDatabase()
  if (!db) throw new Error("Database is unavailable")
  await ensureDatabase(db)
  const now = new Date().toISOString()
  await db.prepare(`INSERT INTO outlook_sessions (id, access_token, refresh_token, expires_at, account_email, account_name, created_at, updated_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(id) DO UPDATE SET access_token = excluded.access_token, refresh_token = excluded.refresh_token,
      expires_at = excluded.expires_at, account_email = excluded.account_email, account_name = excluded.account_name, updated_at = excluded.updated_at`)
    .bind(session.id, session.access_token, session.refresh_token, session.expires_at, session.account_email, session.account_name, now, now).run()
}

export async function deleteSession(id: string): Promise<void> {
  const db = getDatabase()
  if (!db) return
  await ensureDatabase(db)
  await db.prepare("DELETE FROM outlook_sessions WHERE id = ?").bind(id).run()
}

export async function writeLog(log: {
  id: string; campaignId: string; partyCode: string; customerName: string; email: string;
  stage: string; invoiceNumbers: string[]; invoiceCount: number; totalDue: number;
  status: string; errorMessage?: string; sentAt: string; sentBy: string;
}): Promise<void> {
  const db = getDatabase()
  if (!db) throw new Error("Database is unavailable")
  await ensureDatabase(db)
  await db.prepare(`INSERT INTO send_logs (id, campaign_id, party_code, customer_name, email, reminder_stage,
    invoice_numbers, invoice_count, total_due, status, error_message, sent_at, sent_by)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`)
    .bind(log.id, log.campaignId, log.partyCode, log.customerName, log.email, log.stage,
      JSON.stringify(log.invoiceNumbers), log.invoiceCount, log.totalDue, log.status,
      log.errorMessage ?? null, log.sentAt, log.sentBy).run()
}

export async function listLogs(limit = 500): Promise<StoredLog[]> {
  const db = getDatabase()
  if (!db) return []
  await ensureDatabase(db)
  const result = await db.prepare(`SELECT id, party_code, customer_name, email, reminder_stage, invoice_numbers,
    invoice_count, total_due, status, error_message, sent_at FROM send_logs ORDER BY sent_at DESC LIMIT ?`)
    .bind(limit).all<StoredLog>()
  return result.results ?? []
}
