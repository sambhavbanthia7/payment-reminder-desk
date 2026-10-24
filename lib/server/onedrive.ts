import { env } from "cloudflare:workers"
import { getDatabase } from "./database"
import type { DriveFile } from "@/lib/onedrive-sync"

export type { DriveFile }

// App-folder access only: the app sees "Apps/JAL Receivables Desk" in the OneDrive, nothing else.
export const ONEDRIVE_SCOPES = "offline_access User.Read Files.ReadWrite.AppFolder"

export type OneDriveConfig = { clientId: string; clientSecret: string; tenant: string; loginUrl: string; graphUrl: string }
type Link = { access_token: string; refresh_token: string; expires_at: number; account_email: string; account_name: string; item_id: string; item_name: string }

type Key = "ONEDRIVE_CLIENT_ID" | "ONEDRIVE_CLIENT_SECRET" | "ONEDRIVE_TENANT" | "ONEDRIVE_LOGIN_URL" | "ONEDRIVE_GRAPH_URL"
const read = (key: Key) => ((env as unknown as Partial<Record<Key, string>>)[key] ?? "").trim()
const localUrl = (url: string) => /^http:\/\/(127\.0\.0\.1|localhost)(:\d+)?(\/|$)/.test(url)

export function oneDriveConfig(): OneDriveConfig | null {
  const clientId = read("ONEDRIVE_CLIENT_ID"), clientSecret = read("ONEDRIVE_CLIENT_SECRET")
  if (!clientId || !clientSecret) return null
  // Overrides exist only for the practice OneDrive on this computer.
  const loginUrl = read("ONEDRIVE_LOGIN_URL"), graphUrl = read("ONEDRIVE_GRAPH_URL")
  return {
    clientId, clientSecret, tenant: read("ONEDRIVE_TENANT") || "consumers",
    loginUrl: loginUrl && localUrl(loginUrl) ? loginUrl.replace(/\/$/, "") : "https://login.microsoftonline.com",
    graphUrl: graphUrl && localUrl(graphUrl) ? graphUrl.replace(/\/$/, "") : "https://graph.microsoft.com/v1.0",
  }
}

async function db() {
  const database = getDatabase()
  if (!database) throw new Error("Database is unavailable")
  await database.prepare(`CREATE TABLE IF NOT EXISTS onedrive_link (
    id TEXT PRIMARY KEY,
    access_token TEXT NOT NULL,
    refresh_token TEXT NOT NULL,
    expires_at INTEGER NOT NULL,
    account_email TEXT NOT NULL,
    account_name TEXT NOT NULL,
    item_id TEXT NOT NULL,
    item_name TEXT NOT NULL,
    updated_at TEXT NOT NULL
  )`).run()
  return database
}

/** One OneDrive connection for the whole desk, made once by the account owner. */
export async function readLink(): Promise<Link | null> {
  return (await db()).prepare("SELECT access_token, refresh_token, expires_at, account_email, account_name, item_id, item_name FROM onedrive_link WHERE id = 'main'").first<Link>()
}

export async function saveLink(link: Link): Promise<void> {
  await (await db()).prepare(`INSERT INTO onedrive_link (id, access_token, refresh_token, expires_at, account_email, account_name, item_id, item_name, updated_at)
    VALUES ('main', ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(id) DO UPDATE SET access_token = excluded.access_token, refresh_token = excluded.refresh_token, expires_at = excluded.expires_at,
      account_email = excluded.account_email, account_name = excluded.account_name, item_id = excluded.item_id, item_name = excluded.item_name, updated_at = excluded.updated_at`)
    .bind(link.access_token, link.refresh_token, link.expires_at, link.account_email, link.account_name, link.item_id, link.item_name, new Date().toISOString()).run()
}

export async function deleteLink(): Promise<void> {
  await (await db()).prepare("DELETE FROM onedrive_link WHERE id = 'main'").run()
}

export async function requestToken(config: OneDriveConfig, fields: Record<string, string>): Promise<{ access_token: string; refresh_token?: string; expires_in?: number }> {
  const body = new URLSearchParams({ client_id: config.clientId, client_secret: config.clientSecret, ...fields })
  const response = await fetch(`${config.loginUrl}/${encodeURIComponent(config.tenant)}/oauth2/v2.0/token`, { method: "POST", headers: { "content-type": "application/x-www-form-urlencoded" }, body })
  const data = await response.json().catch(() => ({})) as Record<string, unknown>
  if (!response.ok || !data.access_token) throw new Error(String(data.error_description ?? data.error ?? "Microsoft sign-in failed"))
  return data as { access_token: string; refresh_token?: string; expires_in?: number }
}

export class GraphError extends Error {
  status: number
  constructor(status: number, message: string) { super(message); this.status = status }
}

/** Calls Microsoft Graph as the connected account, refreshing the token when it is about to expire. */
export async function graph(path: string, init: RequestInit = {}): Promise<Response> {
  const config = oneDriveConfig()
  if (!config) throw new GraphError(503, "OneDrive is not set up yet")
  let link = await readLink()
  if (!link) throw new GraphError(401, "OneDrive is not connected")
  if (link.expires_at < Date.now() + 60_000) {
    try {
      const token = await requestToken(config, { grant_type: "refresh_token", refresh_token: link.refresh_token, scope: ONEDRIVE_SCOPES })
      link = { ...link, access_token: token.access_token, refresh_token: token.refresh_token ?? link.refresh_token, expires_at: Date.now() + Number(token.expires_in ?? 3600) * 1000 }
      await saveLink(link)
    } catch {
      throw new GraphError(401, "The OneDrive connection has expired. Connect OneDrive again.")
    }
  }
  const headers = new Headers(init.headers)
  headers.set("authorization", `Bearer ${link.access_token}`)
  return fetch(`${config.graphUrl}${path}`, { ...init, headers })
}

const FIELDS = "id,name,eTag,cTag,size,lastModifiedDateTime,lastModifiedBy,file"
type Item = { id: string; name: string; eTag: string; cTag: string; size: number; lastModifiedDateTime: string; lastModifiedBy?: { user?: { displayName?: string } }; file?: unknown }
export const toFile = (item: Item): DriveFile => ({ id: item.id, name: item.name, eTag: item.eTag, cTag: item.cTag, size: item.size, modified: item.lastModifiedDateTime, modifiedBy: item.lastModifiedBy?.user?.displayName ?? "" })

async function json<T>(response: Response, fallback: string): Promise<T> {
  if (response.ok) return response.json() as Promise<T>
  const body = await response.json().catch(() => null) as { error?: { message?: string } } | null
  throw new GraphError(response.status, body?.error?.message ?? fallback)
}

export async function fileInfo(id: string): Promise<DriveFile> {
  return toFile(await json<Item>(await graph(`/me/drive/items/${encodeURIComponent(id)}?$select=${FIELDS}`), "The tracker could not be found in OneDrive"))
}

export async function listTrackers(): Promise<DriveFile[]> {
  const data = await json<{ value: Item[] }>(await graph(`/me/drive/special/approot/children?$select=${FIELDS}&$top=200`), "The OneDrive folder could not be read")
  return data.value.filter((item) => item.file && /\.xlsx$/i.test(item.name)).map(toFile)
}

export async function download(id: string): Promise<ArrayBuffer> {
  const response = await graph(`/me/drive/items/${encodeURIComponent(id)}/content`)
  if (!response.ok) throw new GraphError(response.status, "The tracker could not be downloaded from OneDrive")
  return response.arrayBuffer()
}

const XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

/** Replace the file only if it is still the version this computer last saw (If-Match). */
export async function upload(id: string, bytes: ArrayBuffer, eTag: string): Promise<DriveFile> {
  const response = await graph(`/me/drive/items/${encodeURIComponent(id)}/content`, { method: "PUT", headers: { "content-type": XLSX, ...(eTag ? { "if-match": eTag } : {}) }, body: bytes })
  return toFile(await json<Item>(response, "The tracker could not be saved to OneDrive"))
}

export async function create(name: string, bytes: ArrayBuffer): Promise<DriveFile> {
  const safe = name.replace(/[\\/:*?"<>|#%]/g, "-").trim() || "Debtors Tracker.xlsx"
  const response = await graph(`/me/drive/special/approot:/${encodeURIComponent(/\.xlsx$/i.test(safe) ? safe : `${safe}.xlsx`)}:/content?@microsoft.graph.conflictBehavior=fail`, { method: "PUT", headers: { "content-type": XLSX }, body: bytes })
  return toFile(await json<Item>(response, "The tracker could not be created in OneDrive"))
}

export function errorResponse(error: unknown): Response {
  const status = error instanceof GraphError ? error.status : 500
  return Response.json({ error: error instanceof Error ? error.message : "OneDrive request failed" }, { status })
}
