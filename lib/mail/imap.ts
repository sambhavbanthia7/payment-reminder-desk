import { LineConnection, type MailServer, type OpenSocket } from "./connection.ts"

// Common Sent folder names on cPanel (Dovecot uses "INBOX." as the prefix) and other hosts.
const SENT_NAMES = ["INBOX.Sent", "Sent", "Sent Items", "INBOX.Sent Items", "Sent Messages", "INBOX.Sent Messages"]

function quote(value: string): string {
  if (/[\r\n]/.test(value)) throw new Error("Mailbox settings contain an unsupported character")
  return `"${value.replace(/["\\]/g, "\\$&")}"`
}

/** Parse `* LIST (\HasNoChildren \Sent) "." "INBOX.Sent"` into flags and name. */
export function parseList(line: string): { flags: string[]; name: string } | null {
  const match = /^\* LIST \(([^)]*)\) (?:"(?:[^"\\]|\\.)*"|NIL) (.+)$/i.exec(line)
  if (!match) return null
  const raw = match[2].trim()
  const name = raw.startsWith('"') ? raw.slice(1, -1).replace(/\\(.)/g, "$1") : raw
  return { flags: match[1].split(/\s+/).filter(Boolean), name }
}

export function pickSentFolder(folders: { flags: string[]; name: string }[]): string | null {
  const special = folders.find((f) => f.flags.some((flag) => flag.toLowerCase() === "\\sent"))
  if (special) return special.name
  for (const candidate of SENT_NAMES) {
    const found = folders.find((f) => f.name.toLowerCase() === candidate.toLowerCase())
    if (found) return found.name
  }
  return null
}

/** Store copies of sent messages in the mailbox's Sent folder, so it shows in Outlook. Returns the folder used. */
export async function appendToSent(openSocket: OpenSocket, server: MailServer, messages: string[], folderOverride = ""): Promise<string> {
  const conn = new LineConnection(openSocket(server.host, server.port, server.security))
  let counter = 0
  const run = async (command: string): Promise<string[]> => {
    const tag = `a${++counter}`
    await conn.write(`${tag} ${command}\r\n`)
    return finish(tag)
  }
  const finish = async (tag: string): Promise<string[]> => {
    const lines: string[] = []
    for (;;) {
      const line = await conn.readLine()
      if (line.startsWith(`${tag} `)) {
        if (!/^\S+ OK/i.test(line)) throw new Error(line.slice(tag.length + 1))
        return lines
      }
      lines.push(line)
    }
  }
  try {
    const greeting = await conn.readLine()
    if (!/^\* (OK|PREAUTH)/i.test(greeting)) throw new Error("The mail server did not accept the connection")
    if (server.security === "starttls") { await run("STARTTLS"); conn.upgrade() }
    await run(`LOGIN ${quote(server.user)} ${quote(server.password)}`).catch(() => { throw new Error("The mail server rejected the email address or password") })
    let folder = folderOverride
    if (!folder) {
      const folders = (await run('LIST "" "*"')).map(parseList).filter((f): f is { flags: string[]; name: string } => Boolean(f))
      folder = pickSentFolder(folders) ?? ""
      if (!folder) throw new Error("No Sent folder was found in the mailbox")
    }
    for (const message of messages) {
      const text = message.endsWith("\r\n") ? message : `${message}\r\n`
      const tag = `a${++counter}`
      await conn.write(`${tag} APPEND ${quote(folder)} (\\Seen) {${new TextEncoder().encode(text).length}}\r\n`)
      const ready = await conn.readLine()
      if (!ready.startsWith("+")) throw new Error(`The mail server would not store the copy (${ready})`)
      await conn.write(`${text}\r\n`)
      await finish(tag)
    }
    await run("LOGOUT").catch(() => [])
    return folder
  } finally {
    await conn.close()
  }
}
