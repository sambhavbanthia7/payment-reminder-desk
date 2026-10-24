import { env } from "cloudflare:workers"
import { getDatabase } from "./database"
import { seal, unseal } from "./crypto"
import { cookie, getCookie } from "./http"
import { mailSettings } from "./mail"

/**
 * Sign-in with the mailbox account. The mailbox password is needed later to send email, so it
 * is kept on the server, encrypted with SESSION_SECRET, for as long as the sign-in lasts.
 * The browser only holds a random session id in an HttpOnly cookie.
 */
export const SESSION_COOKIE = "jal_session"
export const SESSION_DAYS = 30
const FAILURE_WINDOW_MS = 15 * 60_000
export const MAX_FAILURES = 5

export type Session = { id: string; email: string; password: string }

function secret(): string {
  const value = ((env as unknown as { SESSION_SECRET?: string }).SESSION_SECRET ?? "").trim()
  if (value.length < 32) throw new Error("SESSION_SECRET is missing or shorter than 32 characters")
  return value
}

async function db() {
  const database = getDatabase()
  if (!database) throw new Error("Database is unavailable")
  await database.prepare(`CREATE TABLE IF NOT EXISTS app_sessions (
    id TEXT PRIMARY KEY,
    email TEXT NOT NULL,
    sealed_password TEXT NOT NULL,
    created_at TEXT NOT NULL,
    expires_at INTEGER NOT NULL
  )`).run()
  await database.prepare("CREATE TABLE IF NOT EXISTS login_failures (at INTEGER NOT NULL)").run()
  return database
}

export async function createSession(email: string, password: string): Promise<string> {
  const id = crypto.randomUUID() + crypto.randomUUID().replace(/-/g, "")
  await (await db()).prepare("INSERT INTO app_sessions (id, email, sealed_password, created_at, expires_at) VALUES (?, ?, ?, ?, ?)")
    .bind(id, email, await seal(secret(), password), new Date().toISOString(), Date.now() + SESSION_DAYS * 86400_000).run()
  return id
}

export async function readSession(id: string | null): Promise<Session | null> {
  if (!id) return null
  const database = await db()
  const row = await database.prepare("SELECT id, email, sealed_password, expires_at FROM app_sessions WHERE id = ?").bind(id).first<{ id: string; email: string; sealed_password: string; expires_at: number }>()
  if (!row) return null
  const password = row.expires_at > Date.now() ? await unseal(secret(), row.sealed_password) : null
  // A sign-in also ends when its account is no longer on the allowed list.
  const allowed = mailSettings().settings?.allowed ?? []
  if (password === null || !allowed.includes(row.email)) { await deleteSession(id); return null }
  return { id: row.id, email: row.email, password }
}

export async function deleteSession(id: string): Promise<void> {
  await (await db()).prepare("DELETE FROM app_sessions WHERE id = ?").bind(id).run()
}

/** Sign everyone out of an account, e.g. after the mail server rejects its password. */
export async function deleteSessionsFor(email: string): Promise<void> {
  await (await db()).prepare("DELETE FROM app_sessions WHERE email = ?").bind(email).run()
}

export const sessionFromRequest = (request: Request) => readSession(getCookie(request, SESSION_COOKIE))

export function sessionFromCookieHeader(header: string): Promise<Session | null> {
  const match = header.split(";").map((part) => part.trim()).find((part) => part.startsWith(`${SESSION_COOKIE}=`))
  return readSession(match ? decodeURIComponent(match.slice(SESSION_COOKIE.length + 1)) : null)
}

export const sessionCookie = (request: Request, id: string) => cookie(request, SESSION_COOKIE, id, SESSION_DAYS * 86400)
export const clearSessionCookie = (request: Request) => cookie(request, SESSION_COOKIE, "", 0)

/** For API routes: the signed-in session, or a 401 response to return. */
export async function requireSession(request: Request): Promise<Session | Response> {
  try {
    const session = await sessionFromRequest(request)
    return session ?? Response.json({ error: "Sign in to continue", signedOut: true }, { status: 401 })
  } catch (error) {
    return Response.json({ error: error instanceof Error ? error.message : "Sign-in is unavailable" }, { status: 503 })
  }
}

// Wrong passwords are counted for the whole app, so repeated guesses cannot get the mailbox
// blocked by the hosting company's brute-force protection.
export async function recentFailures(): Promise<number> {
  const database = await db()
  await database.prepare("DELETE FROM login_failures WHERE at < ?").bind(Date.now() - FAILURE_WINDOW_MS).run()
  const row = await database.prepare("SELECT COUNT(*) AS n FROM login_failures").first<{ n: number }>()
  return Number(row?.n ?? 0)
}

export async function recordFailure(): Promise<void> {
  await (await db()).prepare("INSERT INTO login_failures (at) VALUES (?)").bind(Date.now()).run()
}
