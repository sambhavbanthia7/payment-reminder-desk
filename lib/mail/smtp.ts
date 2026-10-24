import { LineConnection, toBase64, type MailServer, type OpenSocket } from "./connection.ts"

export class SmtpError extends Error {
  code: number
  constructor(code: number, message: string) { super(message); this.code = code }
}

type Reply = { code: number; lines: string[] }

/** One authenticated SMTP connection that can send several messages in turn. */
export class SmtpSession {
  private conn: LineConnection
  private constructor(conn: LineConnection) { this.conn = conn }

  static async open(openSocket: OpenSocket, server: MailServer, heloName: string): Promise<SmtpSession> {
    const conn = new LineConnection(openSocket(server.host, server.port, server.security))
    const session = new SmtpSession(conn)
    try {
      await session.expect(await session.reply(), 220, "The mail server did not accept the connection")
      let features = await session.ehlo(heloName)
      if (server.security === "starttls") {
        if (!features.some((f) => /^STARTTLS\b/i.test(f))) throw new SmtpError(0, "The mail server does not offer STARTTLS on this port. Try port 465 with SSL.")
        await session.command("STARTTLS", 220, "The mail server refused to start encryption")
        conn.upgrade()
        features = await session.ehlo(heloName)
      }
      await session.login(server, features)
      return session
    } catch (error) {
      await conn.close()
      throw error
    }
  }

  private async reply(): Promise<Reply> {
    const lines: string[] = []
    for (;;) {
      const line = await this.conn.readLine()
      lines.push(line.slice(4))
      if (line[3] !== "-") return { code: Number(line.slice(0, 3)), lines }
    }
  }

  private async expect(reply: Reply, code: number, message: string): Promise<Reply> {
    if (Math.floor(reply.code / 100) !== Math.floor(code / 100)) throw new SmtpError(reply.code, `${message} (${reply.code} ${reply.lines.join(" ")})`)
    return reply
  }

  private async command(line: string, code: number, message: string): Promise<Reply> {
    await this.conn.write(`${line}\r\n`)
    return this.expect(await this.reply(), code, message)
  }

  private async ehlo(name: string): Promise<string[]> {
    return (await this.command(`EHLO ${name}`, 250, "The mail server rejected the greeting")).lines
  }

  private async login(server: MailServer, features: string[]): Promise<void> {
    const auth = features.find((f) => /^AUTH\b/i.test(f))?.toUpperCase() ?? ""
    const failed = "The mail server rejected the email address or password"
    if (/\bPLAIN\b/.test(auth) || !auth) {
      await this.command(`AUTH PLAIN ${toBase64(`\u0000${server.user}\u0000${server.password}`)}`, 235, failed)
    } else {
      await this.command("AUTH LOGIN", 334, failed)
      await this.command(toBase64(server.user), 334, failed)
      await this.command(toBase64(server.password), 235, failed)
    }
  }

  /** Send one message. Returns the server's acceptance line, e.g. "250 OK id=…". */
  async send(from: string, recipients: string[], message: string): Promise<string> {
    try {
      await this.command(`MAIL FROM:<${from}>`, 250, "The mail server refused the sender")
      for (const to of recipients) await this.command(`RCPT TO:<${to}>`, 250, `The mail server refused the recipient ${to}`)
      await this.command("DATA", 354, "The mail server refused the message")
      // Dot-stuffing: a line starting with "." gets an extra "." so it cannot end the message early.
      const body = message.replace(/\r?\n/g, "\r\n").replace(/^\./gm, "..")
      await this.conn.write(`${body}${body.endsWith("\r\n") ? "" : "\r\n"}.\r\n`)
      const accepted = await this.expect(await this.reply(), 250, "The mail server did not accept the message")
      return `${accepted.code} ${accepted.lines.join(" ")}`.trim()
    } catch (error) {
      // Clear the half-finished transaction so the next message can still go out on this connection.
      await this.conn.write("RSET\r\n").then(() => this.reply()).catch(() => {})
      throw error
    }
  }

  async close(): Promise<void> {
    await this.conn.write("QUIT\r\n").then(() => this.reply()).catch(() => {})
    await this.conn.close()
  }
}
