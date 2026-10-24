/** The part of a TCP socket the mail clients need. Cloudflare's `connect()` and Node streams both fit. */
export type MailSocket = {
  readable: ReadableStream<Uint8Array>
  writable: WritableStream<Uint8Array>
  close: () => unknown
  /** Upgrade a plain connection after STARTTLS. Returns the encrypted socket. */
  startTls?: () => MailSocket
}
export type Security = "ssl" | "starttls" | "none"
export type OpenSocket = (host: string, port: number, security: Security) => MailSocket
export type MailServer = { host: string; port: number; security: Security; user: string; password: string }

const TIMEOUT_MS = 30_000

/** Line-oriented reader/writer over a socket, used by both SMTP and IMAP. */
export class LineConnection {
  private reader: ReadableStreamDefaultReader<Uint8Array>
  private writer: WritableStreamDefaultWriter<Uint8Array>
  private buffer = new Uint8Array(0)
  private decoder = new TextDecoder("utf-8")

  private socket: MailSocket

  constructor(socket: MailSocket) {
    this.socket = socket
    this.reader = socket.readable.getReader()
    this.writer = socket.writable.getWriter()
  }

  async readLine(): Promise<string> {
    for (;;) {
      const end = this.buffer.findIndex((byte, i) => byte === 10 && this.buffer[i - 1] === 13)
      if (end >= 0) {
        const line = this.decoder.decode(this.buffer.subarray(0, end - 1))
        this.buffer = this.buffer.slice(end + 1)
        return line
      }
      let timer: ReturnType<typeof setTimeout> | undefined
      const timeout = new Promise<never>((_, reject) => { timer = setTimeout(() => reject(new Error("The mail server stopped responding")), TIMEOUT_MS) })
      const { value, done } = await Promise.race([this.reader.read(), timeout]).finally(() => clearTimeout(timer))
      if (done) throw new Error("The mail server closed the connection")
      const next = new Uint8Array(this.buffer.length + value.length)
      next.set(this.buffer); next.set(value, this.buffer.length)
      this.buffer = next
    }
  }

  async write(text: string): Promise<void> {
    await this.writer.write(new TextEncoder().encode(text))
  }

  /** Switch to TLS after the server agreed to STARTTLS. */
  upgrade(): void {
    if (!this.socket.startTls) throw new Error("This connection cannot be encrypted")
    this.reader.releaseLock(); this.writer.releaseLock()
    this.socket = this.socket.startTls()
    this.reader = this.socket.readable.getReader()
    this.writer = this.socket.writable.getWriter()
    this.buffer = new Uint8Array(0)
  }

  async close(): Promise<void> {
    try { this.reader.releaseLock(); this.writer.releaseLock() } catch { /* already released */ }
    try { await this.socket.close() } catch { /* already closed */ }
  }
}

export const toBase64 = (text: string) => {
  const bytes = new TextEncoder().encode(text)
  let binary = ""
  for (const byte of bytes) binary += String.fromCharCode(byte)
  return btoa(binary)
}
