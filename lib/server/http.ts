export function getCookie(request: Request, name: string): string | null {
  const cookie = request.headers.get("cookie") ?? ""
  const match = cookie.split(";").map((part) => part.trim()).find((part) => part.startsWith(`${name}=`))
  return match ? decodeURIComponent(match.slice(name.length + 1)) : null
}

export function cookie(request: Request, name: string, value: string, maxAge = 600): string {
  const url = new URL(request.url)
  // Microsoft permits an http://localhost callback. Keep Secure on every public host.
  const localHttp = url.protocol === "http:" && ["localhost", "127.0.0.1"].includes(url.hostname)
  return `${name}=${encodeURIComponent(value)}; Path=/; HttpOnly; ${localHttp ? "" : "Secure; "}SameSite=Lax; Max-Age=${maxAge}`
}

export function randomUrlSafe(bytes = 32): string {
  return toBase64Url(crypto.getRandomValues(new Uint8Array(bytes)))
}

export async function pkceChallenge(verifier: string): Promise<string> {
  return toBase64Url(new Uint8Array(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier))))
}

function toBase64Url(bytes: Uint8Array): string {
  let binary = ""
  for (const byte of bytes) binary += String.fromCharCode(byte)
  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "")
}

/** Refuse cross-site writes; the dashboard always calls from its own origin. */
export function crossSite(request: Request): Response | null {
  const origin = request.headers.get("origin")
  return origin && origin !== new URL(request.url).origin ? Response.json({ error: "Cross-site request refused" }, { status: 403 }) : null
}
