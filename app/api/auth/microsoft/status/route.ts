import { microsoftConfig, getActiveSession } from "@/lib/server/microsoft"

export async function GET(request: Request) {
  const configured = Boolean(microsoftConfig())
  if (!configured) return Response.json({ configured: false, connected: false })
  try {
    const session = await getActiveSession(request)
    return Response.json({ configured: true, connected: Boolean(session), name: session?.account_name, email: session?.account_email })
  } catch {
    return Response.json({ configured: true, connected: false })
  }
}
