import { connect } from "cloudflare:sockets"
import { env } from "cloudflare:workers"
import type { MailServer, MailSocket, OpenSocket, Security } from "@/lib/mail/connection"

/** Mail server settings. The account itself comes from whoever signed in. */
export type MailSettings = {
  fromName: string; replyTo: string
  /** Hidden copy of every live email, so the team sees it in Outlook whether it uses POP or IMAP. */
  bcc: string[]
  /** Mailbox addresses allowed to sign in to the app. */
  allowed: string[]
  smtp: Omit<MailServer, "user" | "password">
  /** Null when no incoming server is set; sent copies are then not saved to the Sent folder. */
  imap: Omit<MailServer, "user" | "password"> | null
  sentFolder: string
  hourlyLimit: number
}
export type EmailDelivery = { mode: "test" | "live"; testRecipient: string }

type Key = "EMAIL_ALLOWED_ACCOUNTS" | "EMAIL_SMTP_HOST" | "EMAIL_SMTP_PORT" | "EMAIL_SMTP_SECURITY" | "EMAIL_IMAP_HOST" | "EMAIL_IMAP_PORT"
  | "EMAIL_IMAP_SECURITY" | "EMAIL_SENT_FOLDER" | "EMAIL_FROM_NAME" | "EMAIL_REPLY_TO" | "EMAIL_BCC" | "EMAIL_HOURLY_LIMIT" | "EMAIL_MODE" | "EMAIL_TEST_RECIPIENT"
const read = (key: Key) => ((env as unknown as Partial<Record<Key, string>>)[key] ?? "").trim()
const local = (host: string) => ["localhost", "127.0.0.1", "::1"].includes(host.toLowerCase())

function security(value: string, port: number, host: string, implicitPort: number): Security {
  const chosen = value.toLowerCase().replace(/[^a-z]/g, "")
  const result: Security = chosen === "ssl" || chosen === "tls" || chosen === "ssltls" ? "ssl" : chosen === "starttls" ? "starttls" : chosen === "none" ? "none" : port === implicitPort ? "ssl" : "starttls"
  // Unencrypted connections are only for the practice mail server on this computer.
  if (result === "none" && !local(host)) throw new Error("Email must be encrypted. Set the security to SSL or STARTTLS.")
  return result
}

const isEmail = (value: string) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value)
const list = (value: string) => value.split(/[;,\s]+/).map((v) => v.toLowerCase()).filter(Boolean)

/**
 * Reads the mail server settings. Returns `settings: null` with a plain-language `problem` when
 * email is not set up yet or a setting is wrong.
 */
export function mailSettings(): { settings: MailSettings | null; problem: string } {
  const smtpHost = read("EMAIL_SMTP_HOST"), allowed = list(read("EMAIL_ALLOWED_ACCOUNTS"))
  if (!smtpHost && !allowed.length) return { settings: null, problem: "" }
  try {
    if (!smtpHost) throw new Error("EMAIL_SMTP_HOST is missing")
    if (!allowed.length || !allowed.every(isEmail)) throw new Error("EMAIL_ALLOWED_ACCOUNTS must list the mailbox address(es) allowed to sign in")
    const smtpPort = Number(read("EMAIL_SMTP_PORT") || 465)
    if (smtpPort === 25) throw new Error("Port 25 cannot be used. Use 465 (SSL) or 587 (STARTTLS).")
    const imapHost = read("EMAIL_IMAP_HOST"), imapPort = Number(read("EMAIL_IMAP_PORT") || 993)
    const bcc = list(read("EMAIL_BCC"))
    if (!bcc.every(isEmail)) throw new Error("EMAIL_BCC contains an address that is not valid")
    return {
      settings: {
        fromName: read("EMAIL_FROM_NAME") || "JAL Accounts Receivable", replyTo: read("EMAIL_REPLY_TO").toLowerCase(), bcc, allowed,
        smtp: { host: smtpHost, port: smtpPort, security: security(read("EMAIL_SMTP_SECURITY"), smtpPort, smtpHost, 465) },
        imap: imapHost ? { host: imapHost, port: imapPort, security: security(read("EMAIL_IMAP_SECURITY"), imapPort, imapHost, 993) } : null,
        sentFolder: read("EMAIL_SENT_FOLDER"),
        hourlyLimit: Math.max(1, Number(read("EMAIL_HOURLY_LIMIT")) || 100),
      },
      problem: "",
    }
  } catch (error) {
    return { settings: null, problem: error instanceof Error ? error.message : "Email settings are invalid" }
  }
}

/** The mail server settings with the signed-in account's address and password. */
export const withAccount = (server: Omit<MailServer, "user" | "password">, user: string, password: string): MailServer => ({ ...server, user, password })

/**
 * Sending is in test mode unless EMAIL_MODE is exactly "live". Test mode redirects every
 * message to EMAIL_TEST_RECIPIENT, and refuses to send at all when that is not set.
 */
export function emailDelivery(): EmailDelivery {
  return { mode: read("EMAIL_MODE").toLowerCase() === "live" ? "live" : "test", testRecipient: list(read("EMAIL_TEST_RECIPIENT")).join(", ") }
}

export const openSocket: OpenSocket = (host, port, mode) =>
  connect({ hostname: host, port }, { secureTransport: mode === "ssl" ? "on" : mode === "starttls" ? "starttls" : "off", allowHalfOpen: false }) as unknown as MailSocket
