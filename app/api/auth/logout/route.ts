import { crossSite } from "@/lib/server/http"
import { clearSessionCookie, deleteSession, sessionFromRequest } from "@/lib/server/session"

export async function POST(request: Request) {
  const refused = crossSite(request)
  if (refused) return refused
  const session = await sessionFromRequest(request).catch(() => null)
  if (session) await deleteSession(session.id)
  return Response.json({ ok: true }, { headers: { "set-cookie": clearSessionCookie(request) } })
}
