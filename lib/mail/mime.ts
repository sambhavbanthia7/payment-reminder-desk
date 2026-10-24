export type OutgoingMail = {
  fromName: string; from: string; to: string[]; replyTo?: string
  subject: string; text: string; date?: Date; messageId?: string
}

const encoder = new TextEncoder()

function base64(bytes: Uint8Array): string {
  let binary = ""
  for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000))
  return btoa(binary)
}

const clean = (value: string) => value.replace(/[\r\n]+/g, " ").trim()

/** RFC 2047 encoded-words for non-ASCII header text (₹, customer names), split on character boundaries. */
export function encodeHeader(value: string): string {
  const text = clean(value)
  if (/^[\x20-\x7e]*$/.test(text)) return text
  const words: string[] = []
  let chunk = ""
  for (const char of text) {
    if (encoder.encode(chunk + char).length > 45) { words.push(chunk); chunk = "" }
    chunk += char
  }
  if (chunk) words.push(chunk)
  return words.map((w) => `=?UTF-8?B?${base64(encoder.encode(w))}?=`).join("\r\n ")
}

export function formatAddress(name: string, address: string): string {
  const display = clean(name)
  if (!display) return `<${address}>`
  return /^[\x20-\x7e]*$/.test(display) ? `"${display.replace(/["\\]/g, "\\$&")}" <${address}>` : `${encodeHeader(display)} <${address}>`
}

/** A complete RFC 5322 message with CRLF line endings and a base64 UTF-8 text body. */
export function buildMessage(mail: OutgoingMail): string {
  const domain = mail.from.split("@")[1] ?? "localhost"
  const body = base64(encoder.encode(mail.text.replace(/\r?\n/g, "\r\n"))).replace(/.{1,76}/g, "$&\r\n")
  const headers = [
    `From: ${formatAddress(mail.fromName, mail.from)}`,
    `To: ${mail.to.join(", ")}`,
    ...(mail.replyTo ? [`Reply-To: ${mail.replyTo}`] : []),
    `Subject: ${encodeHeader(mail.subject)}`,
    `Date: ${(mail.date ?? new Date()).toUTCString().replace("GMT", "+0000")}`,
    `Message-ID: ${mail.messageId ?? `<${crypto.randomUUID()}@${domain}>`}`,
    "MIME-Version: 1.0",
    "Content-Type: text/plain; charset=UTF-8",
    "Content-Transfer-Encoding: base64",
  ]
  return `${headers.join("\r\n")}\r\n\r\n${body}`
}
