import { requireSession } from "@/lib/server/session"
import { crossSite } from "@/lib/server/http"
import { download, errorResponse, fileInfo, GraphError, readLink, upload } from "@/lib/server/onedrive"

async function chosen(): Promise<string> {
  const link = await readLink()
  if (!link) throw new GraphError(401, "OneDrive is not connected")
  if (!link.item_id) throw new GraphError(404, "No tracker has been chosen in OneDrive yet")
  return link.item_id
}

/** Download the shared tracker, with the version tag the next save must match. */
export async function GET(request: Request) {
  const signedIn = await requireSession(request)
  if (signedIn instanceof Response) return signedIn
  try {
    const id = await chosen()
    const file = await fileInfo(id)
    const bytes = await download(id)
    return new Response(bytes, { headers: { "content-type": "application/octet-stream", "cache-control": "no-store", "x-file": encodeURIComponent(JSON.stringify(file)) } })
  } catch (error) {
    return errorResponse(error)
  }
}

/**
 * Save the workbook over the shared tracker. `x-base-etag`/`x-base-ctag` name the version this
 * computer started from. If someone else changed the file since, nothing is written (409).
 */
export async function PUT(request: Request) {
  const signedIn = await requireSession(request)
  if (signedIn instanceof Response) return signedIn
  const refused = crossSite(request)
  if (refused) return refused
  try {
    const id = await chosen()
    const bytes = await request.arrayBuffer()
    if (!bytes.byteLength || bytes.byteLength > 20 * 1024 * 1024) return Response.json({ error: "The workbook must be between 1 byte and 20 MB" }, { status: 400 })
    const force = request.headers.get("x-force") === "1"
    let eTag = request.headers.get("x-base-etag") ?? ""
    if (force) eTag = (await fileInfo(id)).eTag
    try {
      return Response.json({ file: await upload(id, bytes, eTag) })
    } catch (error) {
      if (!(error instanceof GraphError) || error.status !== 412) throw error
      // The tag also changes for things like a rename. Only a content change is a real conflict.
      const current = await fileInfo(id)
      if (current.cTag !== request.headers.get("x-base-ctag")) return Response.json({ error: "changed", file: current }, { status: 409 })
      return Response.json({ file: await upload(id, bytes, current.eTag) })
    }
  } catch (error) {
    if (error instanceof GraphError && error.status === 423) return Response.json({ error: "locked" }, { status: 423 })
    return errorResponse(error)
  }
}
