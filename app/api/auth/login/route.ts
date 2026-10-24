import { crossSite } from "@/lib/server/http"
import { mailSettings, openSocket, withAccount } from "@/lib/server/mail"
import { createSession, MAX_FAILURES, recentFailures, recordFailure, sessionCookie } from "@/lib/server/session"
import { SmtpError, SmtpSession } from "@/lib/mail/smtp"

/** Sign in with the mailbox address and password. They are checked with the mail server; nothing is sent. */
export async function POST(request: Request) {
  const refused = crossSite(request)
  if (refused) return refused
  const { settings, problem } = mailSettings()
  if (!settings) return Response.json({ error: problem || "Email is not set up on this server yet" }, { status: 503 })
  const body = await request.json().catch(() => ({})) as { email?: string; password?: string }
  const email = String(body.email ?? "").trim().toLowerCase(), password = String(body.password ?? "")
  if (!email || !password) return Response.json({ error: "Enter the email address and password" }, { status: 400 })
  if (!settings.allowed.includes(email)) return Response.json({ error: `This app only accepts ${settings.allowed.join(" or ")}` }, { status: 403 })
  if (await recentFailures() >= MAX_FAILURES) {
    return Response.json({ error: "Too many wrong passwords. Wait 15 minutes before trying again, so the mail server does not block the mailbox." }, { status: 429 })
  }
  try {
    const session = await SmtpSession.open(openSocket, withAccount(settings.smtp, email, password), email.split("@")[1])
    await session.close()
  } catch (error) {
    if (error instanceof SmtpError && error.code >= 500 && error.code < 600) {
      await recordFailure()
      return Response.json({ error: "The email address or password is not correct" }, { status: 401 })
    }
    return Response.json({ error: `Could not reach the mail server: ${error instanceof Error ? error.message : "unknown error"}` }, { status: 502 })
  }
  const id = await createSession(email, password)
  return Response.json({ email }, { headers: { "set-cookie": sessionCookie(request, id) } })
}
