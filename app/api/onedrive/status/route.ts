import { requireSession } from "@/lib/server/session"
import { fileInfo, GraphError, oneDriveConfig, readLink } from "@/lib/server/onedrive"
import type { OneDriveStatus } from "@/lib/onedrive-sync"

export async function GET(request: Request) {
  const signedIn = await requireSession(request)
  if (signedIn instanceof Response) return signedIn
  if (!oneDriveConfig()) return Response.json({ configured: false, connected: false } satisfies OneDriveStatus)
  const link = await readLink()
  if (!link) return Response.json({ configured: true, connected: false } satisfies OneDriveStatus)
  const account = link.account_name && link.account_name !== link.account_email ? `${link.account_name} (${link.account_email})` : link.account_email
  if (!link.item_id) return Response.json({ configured: true, connected: true, account, file: null } satisfies OneDriveStatus)
  try {
    return Response.json({ configured: true, connected: true, account, file: await fileInfo(link.item_id) } satisfies OneDriveStatus)
  } catch (error) {
    const expired = error instanceof GraphError && error.status === 401
    return Response.json({ configured: true, connected: !expired, account, file: null, error: error instanceof Error ? error.message : "OneDrive could not be reached" } satisfies OneDriveStatus)
  }
}
