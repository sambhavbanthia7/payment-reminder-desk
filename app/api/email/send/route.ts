import { claimDraftSend, countAcceptedSince, finishDraftSend } from "@/lib/server/database"
import { emailDelivery, mailSettings, openSocket, withAccount } from "@/lib/server/mail"
import { deleteSessionsFor, requireSession } from "@/lib/server/session"
import { composeMail, envelopeRecipients, validDraft, type DraftPayload, type SendResult } from "@/lib/email-dispatch"
import { buildMessage } from "@/lib/mail/mime"
import { emailList } from "@/lib/debtors"
import { SmtpError, SmtpSession } from "@/lib/mail/smtp"
import { appendToSent } from "@/lib/mail/imap"

const MAX_PER_REQUEST = 10
const PAUSE_MS = 500

export async function POST(request: Request) {
  // Reject cross-site form posts; the dashboard always sends JSON from the same origin.
  const origin = request.headers.get("origin")
  if (origin && origin !== new URL(request.url).origin) return Response.json({ error: "Cross-site request refused" }, { status: 403 })
  if (!request.headers.get("content-type")?.includes("application/json")) return Response.json({ error: "Expected JSON" }, { status: 415 })

  const payload = await request.json().catch(() => null) as { drafts?: DraftPayload[]; senderName?: string } | null
  const drafts = Array.isArray(payload?.drafts) ? payload.drafts : []
  if (!drafts.length) return Response.json({ error: "No drafts were selected" }, { status: 400 })
  if (drafts.length > MAX_PER_REQUEST) return Response.json({ error: `Send at most ${MAX_PER_REQUEST} drafts per request` }, { status: 400 })

  const account = await requireSession(request)
  if (account instanceof Response) return account
  const { settings: config, problem } = mailSettings()
  if (!config) return Response.json({ error: problem || "Email sending is not set up yet" }, { status: 503 })
  const sender = account.email
  const delivery = emailDelivery()
  const testInboxes = emailList(delivery.testRecipient)
  if (delivery.mode === "test" && (!testInboxes.length || !testInboxes.every((v) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v)))) {
    return Response.json({ error: "Test mode is on but EMAIL_TEST_RECIPIENT is not set. No email was sent." }, { status: 503 })
  }
  // The dashboard has no personal logins, so the name is a log label typed on each computer.
  const sentBy = String(payload?.senderName ?? "").replace(/\s+/g, " ").trim().slice(0, 80) || "Unnamed"

  // Stay under the hosting company's hourly limit so the mailbox is not blocked.
  const used = await countAcceptedSince(new Date(Date.now() - 3600_000).toISOString())
  const allowed = Math.max(0, config.hourlyLimit - used)
  if (!allowed) return Response.json({ error: `The hourly sending limit (${config.hourlyLimit}) has been reached. Wait a while and send the rest.` }, { status: 429 })
  const batch = drafts.slice(0, allowed)

  let session: SmtpSession
  try {
    session = await SmtpSession.openWithRetry(openSocket, withAccount(config.smtp, sender, account.password), sender.split("@")[1])
  } catch (error) {
    if (error instanceof SmtpError && error.code >= 500 && error.code < 600) {
      // The mailbox password was changed. Sign everyone out so they sign in with the new one.
      await deleteSessionsFor(sender)
      return Response.json({ error: "The mail server no longer accepts the saved password. Sign in again with the current password.", signedOut: true }, { status: 401 })
    }
    const where = (request as Request & { cf?: { colo?: string } }).cf?.colo ?? "local"
    console.error(`[send] mail server unreachable from ${where} for ${sender}: ${error instanceof Error ? error.message : String(error)}`)
    return Response.json({ error: `Could not reach the mail server (${error instanceof Error ? error.message : "unknown error"}). No email was sent; try again in a minute.` }, { status: 502 })
  }

  const results: SendResult[] = []
  const copies: { index: number; raw: string }[] = []
  try {
    for (const draft of batch) {
      const at = new Date().toISOString()
      if (!validDraft(draft)) {
        results.push({ id: String(draft?.id ?? ""), status: "Failed", sentAt: at, sentTo: "", sentBy, mode: delivery.mode, detail: "Draft is incomplete or has an invalid recipient" })
        continue
      }
      const mail = composeMail(draft, delivery.mode, delivery.testRecipient)
      const sentTo = mail.recipients.join("; ")
      const key = `${delivery.mode}:${draft.id}`
      const earlier = await claimDraftSend(key, sentTo, sentBy)
      if (earlier) {
        // Never send twice. An unfinished earlier attempt may or may not have reached the server.
        const detail = earlier.status === "accepted" ? "Already sent earlier; not sent again" : "An earlier send attempt did not finish. Check the Sent folder before retrying."
        results.push({ id: draft.id, status: earlier.status === "accepted" ? "Accepted" : "Failed", sentAt: earlier.sent_at, sentTo: earlier.sent_to, sentBy: earlier.sent_by, mode: delivery.mode, detail, repeat: true })
        continue
      }
      const raw = buildMessage({ fromName: config.fromName, from: sender, to: mail.recipients, replyTo: config.replyTo || undefined, subject: mail.subject, text: mail.text })
      try {
        await session.send(sender, envelopeRecipients(mail.recipients, config.bcc, delivery.mode), raw)
        // The server accepting the message is not proof the customer received it; bounces arrive in the mailbox.
        const detail = `Accepted by the mail server${delivery.mode === "live" && config.bcc.length ? ` · copy sent to ${config.bcc.join(", ")}` : ""}`
        await finishDraftSend(key, "accepted", detail)
        results.push({ id: draft.id, status: "Accepted", sentAt: new Date().toISOString(), sentTo, sentBy, mode: delivery.mode, detail })
        copies.push({ index: results.length - 1, raw })
      } catch (error) {
        const detail = error instanceof Error ? error.message : "The mail server did not accept the message"
        await finishDraftSend(key, "failed", detail)
        results.push({ id: draft.id, status: "Failed", sentAt: new Date().toISOString(), sentTo, sentBy, mode: delivery.mode, detail })
      }
      await new Promise((resolve) => setTimeout(resolve, PAUSE_MS))
    }
  } finally {
    await session.close()
  }

  // Emails sent by an app do not appear in Outlook's Sent folder by themselves, so file a copy there.
  if (config.imap && copies.length) {
    try {
      const folder = await appendToSent(openSocket, withAccount(config.imap, sender, account.password), copies.map((c) => c.raw), config.sentFolder)
      for (const { index } of copies) results[index].detail += ` · copy saved in ${folder}`
    } catch (error) {
      for (const { index } of copies) results[index].detail += ` · copy not saved in the Sent folder (${error instanceof Error ? error.message : "unknown error"})`
    }
  }
  return Response.json({ results, mode: delivery.mode, skipped: drafts.length - batch.length })
}
