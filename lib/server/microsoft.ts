import { readSession, writeSession, type StoredSession } from "./database"

export type MicrosoftConfig = {
  clientId: string
  clientSecret: string | null
  tenantId: string
}

export function microsoftConfig(): MicrosoftConfig | null {
  const clientId = process.env.MICROSOFT_CLIENT_ID?.trim()
  if (!clientId) return null
  return {
    clientId,
    clientSecret: process.env.MICROSOFT_CLIENT_SECRET?.trim() || null,
    tenantId: process.env.MICROSOFT_TENANT_ID?.trim() || "organizations",
  }
}

export function getCookie(request: Request, name: string): string | null {
  const cookie = request.headers.get("cookie") ?? ""
  const match = cookie.split(";").map((part) => part.trim()).find((part) => part.startsWith(`${name}=`))
  return match ? decodeURIComponent(match.slice(name.length + 1)) : null
}

export function cookie(name: string, value: string, maxAge = 3600): string {
  return `${name}=${encodeURIComponent(value)}; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=${maxAge}`
}

export async function getActiveSession(request: Request): Promise<StoredSession | null> {
  const id = getCookie(request, "payment_outlook_session")
  if (!id) return null
  const session = await readSession(id)
  if (!session) return null
  if (session.expires_at > Date.now() + 60_000) return session
  if (!session.refresh_token) return null
  const config = microsoftConfig()
  if (!config) return null

  const token = await requestToken(config, {
    grant_type: "refresh_token",
    refresh_token: session.refresh_token,
    scope: "openid profile offline_access User.Read Mail.Send",
  })
  const refreshed = {
    ...session,
    access_token: String(token.access_token),
    refresh_token: String(token.refresh_token ?? session.refresh_token),
    expires_at: Date.now() + Number(token.expires_in ?? 3600) * 1000,
  }
  await writeSession(refreshed)
  return refreshed
}

export async function requestToken(config: MicrosoftConfig, fields: Record<string, string>): Promise<Record<string, unknown>> {
  const body = new URLSearchParams({ client_id: config.clientId, ...fields })
  if (config.clientSecret) body.set("client_secret", config.clientSecret)
  const response = await fetch(`https://login.microsoftonline.com/${encodeURIComponent(config.tenantId)}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body,
  })
  const data = await response.json() as Record<string, unknown>
  if (!response.ok) throw new Error(String(data.error_description ?? data.error ?? "Microsoft token exchange failed"))
  return data
}

export function randomUrlSafe(bytes = 32): string {
  const buffer = crypto.getRandomValues(new Uint8Array(bytes))
  return toBase64Url(buffer)
}

export async function pkceChallenge(verifier: string): Promise<string> {
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier))
  return toBase64Url(new Uint8Array(digest))
}

function toBase64Url(bytes: Uint8Array): string {
  let binary = ""
  for (const byte of bytes) binary += String.fromCharCode(byte)
  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "")
}
