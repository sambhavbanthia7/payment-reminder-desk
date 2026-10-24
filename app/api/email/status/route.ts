import { emailDelivery, mailSettings } from "@/lib/server/mail"
import { requireSession } from "@/lib/server/session"
import type { EmailStatus } from "@/lib/email-dispatch"

export async function GET(request: Request) {
  const session = await requireSession(request)
  if (session instanceof Response) return session
  const { settings, problem } = mailSettings()
  if (!settings) return Response.json({ configured: false, ...(problem ? { problem } : {}) } satisfies EmailStatus)
  return Response.json({ configured: true, sender: session.email, ...emailDelivery(), sentCopy: Boolean(settings.imap), copyTo: settings.bcc, hourlyLimit: settings.hourlyLimit } satisfies EmailStatus)
}
