/** AES-GCM encryption for small secrets kept on the server (the mailbox password of a sign-in). */
const encoder = new TextEncoder(), decoder = new TextDecoder()

async function keyFor(secret: string): Promise<CryptoKey> {
  const raw = await crypto.subtle.digest("SHA-256", encoder.encode(secret))
  return crypto.subtle.importKey("raw", raw, "AES-GCM", false, ["encrypt", "decrypt"])
}

const b64 = (bytes: Uint8Array) => { let s = ""; for (const b of bytes) s += String.fromCharCode(b); return btoa(s) }
const unb64 = (text: string) => Uint8Array.from(atob(text), (c) => c.charCodeAt(0))

export async function seal(secret: string, text: string): Promise<string> {
  const iv = crypto.getRandomValues(new Uint8Array(12))
  const data = new Uint8Array(await crypto.subtle.encrypt({ name: "AES-GCM", iv }, await keyFor(secret), encoder.encode(text)))
  return `${b64(iv)}.${b64(data)}`
}

/** Returns null if the value was tampered with or sealed with a different secret. */
export async function unseal(secret: string, value: string): Promise<string | null> {
  try {
    const [iv, data] = value.split(".")
    return decoder.decode(await crypto.subtle.decrypt({ name: "AES-GCM", iv: unb64(iv) }, await keyFor(secret), unb64(data)))
  } catch {
    return null
  }
}
