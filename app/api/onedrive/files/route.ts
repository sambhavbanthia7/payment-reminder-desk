import { requireSession } from "@/lib/server/session"
import { crossSite } from "@/lib/server/http"
import { create, errorResponse, listTrackers, readLink, saveLink } from "@/lib/server/onedrive"

/** Excel files in the desk's OneDrive folder. */
export async function GET(request: Request) {
  const signedIn = await requireSession(request)
  if (signedIn instanceof Response) return signedIn
  try { return Response.json({ files: await listTrackers() }) } catch (error) { return errorResponse(error) }
}

/** Put the workbook open in the browser into OneDrive as a new tracker and use it from now on. */
export async function POST(request: Request) {
  const signedIn = await requireSession(request)
  if (signedIn instanceof Response) return signedIn
  const refused = crossSite(request)
  if (refused) return refused
  try {
    const name = new URL(request.url).searchParams.get("name") ?? "Debtors Tracker.xlsx"
    const bytes = await request.arrayBuffer()
    if (!bytes.byteLength || bytes.byteLength > 20 * 1024 * 1024) return Response.json({ error: "The workbook must be between 1 byte and 20 MB" }, { status: 400 })
    const file = await create(name, bytes)
    const link = await readLink()
    if (link) await saveLink({ ...link, item_id: file.id, item_name: file.name })
    return Response.json({ file })
  } catch (error) {
    return errorResponse(error)
  }
}
