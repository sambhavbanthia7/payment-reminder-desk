import assert from "node:assert/strict"
import net from "node:net"
import { Duplex } from "node:stream"
import test from "node:test"
import { buildMessage, encodeHeader } from "../lib/mail/mime.ts"
import { MailConnectionError, SmtpSession, fallbackFor } from "../lib/mail/smtp.ts"
import { appendToSent, parseList, pickSentFolder } from "../lib/mail/imap.ts"
import { startCatcher } from "../scripts/mail-catcher.mjs"

const openSocket = (host, port) => {
  const socket = net.connect(port, host)
  const { readable, writable } = Duplex.toWeb(socket)
  return { readable, writable, close: () => socket.end() }
}
const decodeBody = (raw) => Buffer.from(raw.split("\r\n\r\n")[1].replace(/\r\n/g, ""), "base64").toString("utf8")
const server = (port, password = "secret") => ({ host: "127.0.0.1", port, security: "none", user: "accounts@jal.example", password })
const mail = (subject, text) => buildMessage({ fromName: "JAL Accounts", from: "accounts@jal.example", to: ["a@customer.example"], subject, text })

test("headers with ₹ and non-English names are encoded and decode back exactly", () => {
  const encoded = encodeHeader("Payment reminder: AGRAWAL SANITATION · ₹10,10,765.00")
  assert.match(encoded, /^=\?UTF-8\?B\?/)
  const decoded = encoded.split("\r\n ").map((w) => Buffer.from(w.slice(10, -2), "base64").toString("utf8")).join("")
  assert.equal(decoded, "Payment reminder: AGRAWAL SANITATION · ₹10,10,765.00")
  assert.equal(encodeHeader("Plain\r\nInjected: header"), "Plain Injected: header")
})

test("SMTP sends several messages on one connection with the right envelope and body", async (t) => {
  const catcher = await startCatcher({ password: "secret" }); t.after(() => catcher.close())
  const session = await SmtpSession.open(openSocket, server(catcher.smtpPort), "jal.example")
  const reply = await session.send("accounts@jal.example", ["a@customer.example", "b@customer.example"], mail("Reminder ₹100", "Dear customer,\n\nTotal ₹100"))
  await session.send("accounts@jal.example", ["c@customer.example"], mail("Second", "Two"))
  await session.close()
  assert.match(reply, /^250 /)
  assert.equal(catcher.messages.length, 2)
  assert.deepEqual(catcher.messages[0].to, ["a@customer.example", "b@customer.example"])
  assert.equal(decodeBody(catcher.messages[0].raw), "Dear customer,\r\n\r\nTotal ₹100")
  assert.match(catcher.messages[0].raw, /^From: "JAL Accounts" <accounts@jal\.example>/m)
})

test("a BCC copy is delivered but never shown in the message headers", async (t) => {
  const catcher = await startCatcher(); t.after(() => catcher.close())
  const session = await SmtpSession.open(openSocket, server(catcher.smtpPort), "jal.example")
  await session.send("accounts@jal.example", ["a@customer.example", "finance@jal.example"], mail("Reminder", "Hello"))
  await session.close()
  assert.deepEqual(catcher.messages[0].to, ["a@customer.example", "finance@jal.example"])
  assert.match(catcher.messages[0].raw, /^To: a@customer\.example$/m)
  assert.doesNotMatch(catcher.messages[0].raw, /finance@jal\.example/)
})

test("lines starting with a dot survive the SMTP end-of-message marker", async (t) => {
  const catcher = await startCatcher(); t.after(() => catcher.close())
  const session = await SmtpSession.open(openSocket, server(catcher.smtpPort), "jal.example")
  await session.send("accounts@jal.example", ["a@customer.example"], "Subject: x\r\n\r\n.\r\n..two\r\nend")
  await session.close()
  assert.equal(catcher.messages[0].raw, "Subject: x\r\n\r\n.\r\n..two\r\nend\r\n")
})

test("a wrong password gives a plain-language error", async (t) => {
  const catcher = await startCatcher({ password: "secret" }); t.after(() => catcher.close())
  await assert.rejects(SmtpSession.open(openSocket, server(catcher.smtpPort, "wrong"), "jal.example"), /rejected the email address or password/)
})

test("a refused recipient fails only that message; the connection keeps working", async (t) => {
  const catcher = await startCatcher({ rejectRecipient: /bounce/ }); t.after(() => catcher.close())
  const session = await SmtpSession.open(openSocket, server(catcher.smtpPort), "jal.example")
  await assert.rejects(session.send("accounts@jal.example", ["bounce@customer.example"], mail("x", "x")), /refused the recipient bounce@customer\.example/)
  await session.send("accounts@jal.example", ["ok@customer.example"], mail("y", "y"))
  await session.close()
  assert.equal(catcher.messages.length, 1)
})

test("a copy is stored in the mailbox's Sent folder over IMAP", async (t) => {
  const catcher = await startCatcher({ password: "secret" }); t.after(() => catcher.close())
  const raw = mail("Reminder ₹100", "Copy for Outlook")
  const folder = await appendToSent(openSocket, server(catcher.imapPort), [raw, mail("Second", "Two")])
  assert.equal(folder, "INBOX.Sent")
  assert.equal(catcher.sentCopies.length, 2)
  assert.equal(catcher.sentCopies[0].raw, raw.endsWith("\r\n") ? raw : `${raw}\r\n`)
  await assert.rejects(appendToSent(openSocket, server(catcher.imapPort, "wrong"), [raw]), /rejected the email address or password/)
})

test("the Sent folder is found by its flag or by common cPanel names", () => {
  assert.deepEqual(parseList('* LIST (\\HasNoChildren \\Sent) "." "INBOX.Sent"'), { flags: ["\\HasNoChildren", "\\Sent"], name: "INBOX.Sent" })
  assert.equal(pickSentFolder([{ flags: [], name: "INBOX" }, { flags: [], name: "INBOX.Sent" }]), "INBOX.Sent")
  assert.equal(pickSentFolder([{ flags: ["\\Sent"], name: "Sent Items" }, { flags: [], name: "INBOX.Sent" }]), "Sent Items")
  assert.equal(pickSentFolder([{ flags: [], name: "INBOX" }]), null)
})

test("a dropped connection is retried, and a wrong password is not", async (t) => {
  const catcher = await startCatcher({ password: "secret", dropFirst: 1 }); t.after(() => catcher.close())
  const session = await SmtpSession.openWithRetry(openSocket, server(catcher.smtpPort), "jal.example", null)
  await session.close()
  const failing = await startCatcher({ password: "secret" }); t.after(() => failing.close())
  await assert.rejects(SmtpSession.openWithRetry(openSocket, server(failing.smtpPort, "wrong"), "jal.example", null), /rejected the email address or password/)
})

test("when one port keeps dropping, the other port is tried", async (t) => {
  const dead = await startCatcher({ dropFirst: 99 }); t.after(() => dead.close())
  const good = await startCatcher(); t.after(() => good.close())
  const session = await SmtpSession.openWithRetry(openSocket, server(dead.smtpPort), "jal.example", server(good.smtpPort))
  await session.close()
  await assert.rejects(SmtpSession.openWithRetry(openSocket, server(dead.smtpPort), "jal.example", null), (e) => e instanceof MailConnectionError && /kept dropping/.test(e.message))
  assert.deepEqual(fallbackFor({ host: "h", port: 465, security: "ssl", user: "u", password: "p" }), { host: "h", port: 587, security: "starttls", user: "u", password: "p" })
})
