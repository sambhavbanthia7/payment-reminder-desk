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

export function getDatabase(): Database | null {
  return ((env as unknown as { DB?: Database }).DB ?? null)
}

export type DraftSend = { id: string; status: "sending" | "accepted" | "failed"; sent_to: string; sent_at: string; sent_by: string; detail: string }

async function ensureDraftSends(db: Database): Promise<void> {
  await db.prepare(`CREATE TABLE IF NOT EXISTS draft_sends (
    id TEXT PRIMARY KEY,
    status TEXT NOT NULL,
    sent_to TEXT NOT NULL,
    sent_at TEXT NOT NULL,
    sent_by TEXT NOT NULL,
    detail TEXT NOT NULL
  )`).run()
}

/**
 * Reserve a draft before contacting the mail server so the same draft can never be sent twice.
 * Returns null when this request owns the send, or the earlier record when it does not.
 * Only a failed attempt may be retried.
 */
export async function claimDraftSend(id: string, sentTo: string, sentBy: string): Promise<DraftSend | null> {
  const db = getDatabase()
  if (!db) throw new Error("Database is unavailable")
  await ensureDraftSends(db)
  const now = new Date().toISOString()
  const result = await db.prepare(`INSERT INTO draft_sends (id, status, sent_to, sent_at, sent_by, detail) VALUES (?, 'sending', ?, ?, ?, '')
    ON CONFLICT(id) DO UPDATE SET status = 'sending', sent_to = excluded.sent_to, sent_at = excluded.sent_at, sent_by = excluded.sent_by, detail = ''
    WHERE draft_sends.status = 'failed'`)
    .bind(id, sentTo, now, sentBy).run() as { meta?: { changes?: number } }
  if ((result.meta?.changes ?? 0) > 0) return null
  return await db.prepare("SELECT id, status, sent_to, sent_at, sent_by, detail FROM draft_sends WHERE id = ?").bind(id).first<DraftSend>()
}

export async function finishDraftSend(id: string, status: "accepted" | "failed", detail: string): Promise<void> {
  const db = getDatabase()
  if (!db) throw new Error("Database is unavailable")
  await db.prepare("UPDATE draft_sends SET status = ?, detail = ?, sent_at = ? WHERE id = ?").bind(status, detail, new Date().toISOString(), id).run()
}

/** Emails the mail server accepted since `since` (ISO time), test and live together, for the hourly limit. */
export async function countAcceptedSince(since: string): Promise<number> {
  const db = getDatabase()
  if (!db) throw new Error("Database is unavailable")
  await ensureDraftSends(db)
  const row = await db.prepare("SELECT COUNT(*) AS n FROM draft_sends WHERE status = 'accepted' AND sent_at > ?").bind(since).first<{ n: number }>()
  return Number(row?.n ?? 0)
}
