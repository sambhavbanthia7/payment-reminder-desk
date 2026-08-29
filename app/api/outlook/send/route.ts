import { writeLog } from "@/lib/server/database"
import { getActiveSession } from "@/lib/server/microsoft"

type Message = {
  campaignId: string
  partyCode: string
  customerName: string
  email: string
  subject: string
  html: string
  stage: string
  invoiceNumbers: string[]
  invoiceCount: number
  totalDue: number
}

export async function POST(request: Request) {
  const payload = await request.json() as { messages?: Message[]; demo?: boolean }
  const messages = Array.isArray(payload.messages) ? payload.messages.slice(0, 50) : []
  if (!messages.length) return Response.json({ error: "No messages were selected" }, { status: 400 })
  const session = payload.demo ? null : await getActiveSession(request)
  if (!payload.demo && !session) return Response.json({ error: "Please reconnect Outlook before sending" }, { status: 401 })
  const logs = []

  for (const message of messages) {
    if (!isMessageValid(message)) continue
    const sentAt = new Date().toISOString()
    let status: "sent" | "failed" | "previewed" = payload.demo ? "previewed" : "sent"
    let errorMessage: string | undefined
    if (!payload.demo) {
      try {
        const response = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
          method: "POST",
          headers: { authorization: `Bearer ${session!.access_token}`, "content-type": "application/json" },
          body: JSON.stringify({ message: { subject: message.subject, body: { contentType: "HTML", content: message.html }, toRecipients: [{ emailAddress: { address: message.email } }] }, saveToSentItems: true }),
        })
        if (!response.ok) throw new Error(`Outlook returned ${response.status}`)
      } catch (error) {
        status = "failed"
        errorMessage = error instanceof Error ? error.message : "Outlook send failed"
      }
    }
    const log = { id: crypto.randomUUID(), ...message, status, errorMessage, sentAt }
    await writeLog({ ...log, sentBy: session?.account_email ?? "sample-preview" })
    logs.push(log)
  }
  return Response.json({ logs })
}

function isMessageValid(message: Message): boolean {
  return Boolean(message.partyCode && message.customerName && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(message.email) && message.subject && message.html && Array.isArray(message.invoiceNumbers) && Number.isFinite(message.totalDue))
}
