import { requireSession } from "@/lib/server/session"
import { crossSite } from "@/lib/server/http"
import { errorResponse, fileInfo, readLink, saveLink } from "@/lib/server/onedrive"

export async function POST(request: Request) {
  const signedIn = await requireSession(request)
  if (signedIn instanceof Response) return signedIn
  const refused = crossSite(request)
  if (refused) return refused
  try {
    const { id } = await request.json() as { id?: string }
    if (!id) return Response.json({ error: "Choose a file" }, { status: 400 })
    const file = await fileInfo(id)
    if (!/\.xlsx$/i.test(file.name)) return Response.json({ error: "Choose an .xlsx Excel file" }, { status: 400 })
    const link = await readLink()
    if (!link) return Response.json({ error: "OneDrive is not connected" }, { status: 401 })
    await saveLink({ ...link, item_id: file.id, item_name: file.name })
    return Response.json({ file })
  } catch (error) {
    return errorResponse(error)
  }
}
