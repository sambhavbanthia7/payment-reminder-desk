// A practice mail server for local testing. It speaks just enough SMTP and IMAP for the
// Receivables Desk: it accepts every message and keeps it instead of delivering it.
//   npm run mail:catcher      → SMTP on 127.0.0.1:2525, IMAP on 127.0.0.1:1143, copies in .mail-catcher/
import net from "node:net"
import fs from "node:fs"
import path from "node:path"
import { fileURLToPath } from "node:url"

function lineReader(socket) {
  let buffer = Buffer.alloc(0)
  const waiting = []
  const pump = () => {
    while (waiting.length) {
      const next = waiting[0]
      if (next.bytes !== undefined) {
        if (buffer.length < next.bytes) return
        waiting.shift().resolve(buffer.subarray(0, next.bytes).toString("utf8")); buffer = buffer.subarray(next.bytes)
      } else {
        const end = buffer.indexOf("\r\n")
        if (end < 0) return
        waiting.shift().resolve(buffer.subarray(0, end).toString("utf8")); buffer = buffer.subarray(end + 2)
      }
    }
  }
  socket.on("data", (chunk) => { buffer = Buffer.concat([buffer, chunk]); pump() })
  return {
    line: () => new Promise((resolve) => { waiting.push({ resolve }); pump() }),
    bytes: (bytes) => new Promise((resolve) => { waiting.push({ bytes, resolve }); pump() }),
  }
}

function decodeAuthPlain(value) {
  const [, user, password] = Buffer.from(value, "base64").toString("utf8").split("\u0000")
  return { user, password }
}

export async function startCatcher({ smtpPort = 0, imapPort = 0, dir = "", password = "", rejectRecipient = null, dropFirst = 0, log = () => {} } = {}) {
  const messages = [], sentCopies = []
  const save = (kind, text) => {
    if (!dir) return
    fs.mkdirSync(path.join(dir, kind), { recursive: true })
    fs.writeFileSync(path.join(dir, kind, `${new Date().toISOString().replace(/[:.]/g, "-")}-${Math.random().toString(36).slice(2, 6)}.eml`), text)
  }
  const authOk = (pw) => !password || pw === password

  let connections = 0
  const smtp = net.createServer((socket) => {
    // dropFirst: cut the first N connections straight away, to test how the app copes with a dropped line.
    if (++connections <= dropFirst) { socket.destroy(); return }
    const read = lineReader(socket)
    const say = (text) => socket.write(`${text}\r\n`)
    let envelope = { from: "", to: [] }, authed = false
    ;(async () => {
      say("220 mail-catcher ready")
      for (;;) {
        const line = await read.line()
        const verb = line.split(" ")[0].toUpperCase()
        if (verb === "EHLO" || verb === "HELO") { say("250-mail-catcher"); say("250-AUTH PLAIN LOGIN"); say("250 SIZE 10485760") }
        else if (verb === "AUTH") {
          const [, mech, initial] = line.split(" ")
          let pw
          if (mech.toUpperCase() === "PLAIN") pw = decodeAuthPlain(initial ?? await (say("334 "), read.line())).password
          else { say("334 VXNlcm5hbWU6"); await read.line(); say("334 UGFzc3dvcmQ6"); pw = Buffer.from(await read.line(), "base64").toString("utf8") }
          if (authOk(pw)) { authed = true; say("235 2.7.0 Authentication successful") } else say("535 5.7.8 Authentication failed")
        }
        else if (verb === "MAIL") { if (!authed) { say("530 5.7.0 Authentication required"); continue } envelope = { from: line.match(/<(.*)>/)?.[1] ?? "", to: [] }; say("250 OK") }
        else if (verb === "RCPT") { const to = line.match(/<(.*)>/)?.[1] ?? ""; if (rejectRecipient && rejectRecipient.test(to)) say("550 5.1.1 No such user"); else { envelope.to.push(to); say("250 OK") } }
        else if (verb === "DATA") {
          say("354 End data with <CR><LF>.<CR><LF>")
          const lines = []
          for (let l = await read.line(); l !== "."; l = await read.line()) lines.push(l.startsWith("..") ? l.slice(1) : l)
          const raw = lines.join("\r\n") + "\r\n"
          messages.push({ ...envelope, raw }); save("outbox", raw)
          log(`[smtp] ${envelope.from} → ${envelope.to.join(", ")} · ${raw.match(/^Subject: (.*)$/m)?.[1] ?? ""}`)
          say(`250 OK queued as catcher-${messages.length}`)
        }
        else if (verb === "RSET" || verb === "NOOP") say("250 OK")
        else if (verb === "QUIT") { say("221 Bye"); socket.end(); return }
        else say("502 Command not implemented")
      }
    })().catch(() => socket.destroy())
  })

  const imap = net.createServer((socket) => {
    const read = lineReader(socket)
    const say = (text) => socket.write(`${text}\r\n`)
    ;(async () => {
      say("* OK [CAPABILITY IMAP4rev1 AUTH=PLAIN] mail-catcher ready")
      for (;;) {
        const line = await read.line()
        const [tag, verbRaw] = line.split(" ")
        const verb = (verbRaw ?? "").toUpperCase()
        if (verb === "LOGIN") { const pw = [...line.matchAll(/"((?:[^"\\]|\\.)*)"/g)].map((m) => m[1].replace(/\\(.)/g, "$1"))[1]; say(authOk(pw) ? `${tag} OK LOGIN completed` : `${tag} NO [AUTHENTICATIONFAILED] Authentication failed`) }
        else if (verb === "LIST") { say('* LIST (\\HasNoChildren) "." "INBOX"'); say('* LIST (\\HasNoChildren \\Sent) "." "INBOX.Sent"'); say('* LIST (\\HasNoChildren \\Trash) "." "INBOX.Trash"'); say(`${tag} OK LIST completed`) }
        else if (verb === "APPEND") {
          const size = Number(line.match(/\{(\d+)\}$/)?.[1])
          const folder = line.match(/APPEND "((?:[^"\\]|\\.)*)"/i)?.[1]
          say("+ Ready for literal data")
          const raw = await read.bytes(size); await read.line()
          sentCopies.push({ folder, raw }); save("sent-folder", raw)
          log(`[imap] copy stored in ${folder}`)
          say(`${tag} OK APPEND completed`)
        }
        else if (verb === "LOGOUT") { say("* BYE logging out"); say(`${tag} OK LOGOUT completed`); socket.end(); return }
        else say(`${tag} BAD Command not implemented`)
      }
    })().catch(() => socket.destroy())
  })

  const listen = (server, port) => new Promise((resolve) => server.listen(port, "127.0.0.1", () => resolve(server.address().port)))
  return {
    smtpPort: await listen(smtp, smtpPort), imapPort: await listen(imap, imapPort), messages, sentCopies,
    close: () => Promise.all([new Promise((r) => smtp.close(r)), new Promise((r) => imap.close(r))]),
  }
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const dir = path.resolve(".mail-catcher")
  // MAIL_CATCHER_PASSWORD makes the practice server check the password, to try wrong-password handling.
  const catcher = await startCatcher({ smtpPort: 2525, imapPort: 1143, dir, password: process.env.MAIL_CATCHER_PASSWORD ?? "", log: (line) => console.log(line) })
  console.log(`Practice mail server running. SMTP 127.0.0.1:${catcher.smtpPort} · IMAP 127.0.0.1:${catcher.imapPort} · emails saved in ${dir}`)
}
