import { requireSession } from "@/lib/server/session"
import { crossSite } from "@/lib/server/http"
import { deleteLink } from "@/lib/server/onedrive"

export async function POST(request: Request) {
  const signedIn = await requireSession(request)
  if (signedIn instanceof Response) return signedIn
  const refused = crossSite(request)
  if (refused) return refused
  await deleteLink()
  return Response.json({ ok: true })
}
