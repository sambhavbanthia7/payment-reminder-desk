/** Browser side of OneDrive sync: status types, the open/poll rules, and calls to /api/onedrive. */

export type DriveFile = { id: string; name: string; eTag: string; cTag: string; modified: string; modifiedBy: string; size: number }
export type OneDriveStatus = { configured: boolean; connected: boolean; account?: string; file?: DriveFile | null; error?: string }
/** The OneDrive version this computer's workbook started from, and the last revision saved there. */
export type DriveBase = { itemId: string; eTag: string; cTag: string; savedRevision: number }

/**
 * On opening the app with a OneDrive tracker chosen:
 * - "use-local": this computer already has that exact version (plus perhaps unsaved edits to push).
 * - "load-remote": take OneDrive's version; nothing on this computer would be lost.
 * - "ask": OneDrive has a different version and this computer has unsaved edits.
 */
export function openDecision(local: { revision: number; exportedRevision: number; drive?: DriveBase } | null, file: DriveFile): "use-local" | "load-remote" | "ask" {
  if (!local) return "load-remote"
  const sameItem = local.drive?.itemId === file.id
  if (sameItem && local.drive?.cTag === file.cTag) return "use-local"
  const saved = sameItem ? local.drive!.savedRevision : local.exportedRevision
  return local.revision > saved ? "ask" : "load-remote"
}

/** While working: has someone else changed the file, and can we simply take their version? */
export function pollDecision(base: DriveBase, revision: number, file: DriveFile): "same" | "reload" | "conflict" {
  if (file.id !== base.itemId) return "reload"
  // The eTag also changes for renames and similar; the cTag changes only when the content does.
  if (file.cTag === base.cTag) return "same"
  return revision > base.savedRevision ? "conflict" : "reload"
}

export class DriveConflict extends Error { file: DriveFile; constructor(file: DriveFile) { super("The tracker was changed in OneDrive by someone else"); this.file = file } }
export class DriveLocked extends Error { constructor() { super("The tracker is open for editing somewhere else. Saving will retry.") } }

async function failure(response: Response, fallback: string): Promise<never> {
  const data = await response.json().catch(() => ({})) as { error?: string; file?: DriveFile }
  if (response.status === 409 && data.file) throw new DriveConflict(data.file)
  if (response.status === 423) throw new DriveLocked()
  throw new Error(data.error ?? fallback)
}

export async function fetchDriveStatus(): Promise<OneDriveStatus> {
  const response = await fetch("/api/onedrive/status", { cache: "no-store" })
  if (!response.ok) return { configured: false, connected: false, error: "OneDrive status unavailable" }
  return response.json() as Promise<OneDriveStatus>
}

export async function downloadTracker(): Promise<{ bytes: ArrayBuffer; file: DriveFile }> {
  const response = await fetch("/api/onedrive/workbook", { cache: "no-store" })
  if (!response.ok) await failure(response, "The tracker could not be opened from OneDrive")
  return { bytes: await response.arrayBuffer(), file: JSON.parse(decodeURIComponent(response.headers.get("x-file") ?? "")) as DriveFile }
}

export async function saveTracker(bytes: Uint8Array, base: DriveBase, force = false): Promise<DriveFile> {
  const response = await fetch("/api/onedrive/workbook", {
    method: "PUT", body: new Blob([bytes.slice().buffer as ArrayBuffer]),
    headers: { "content-type": "application/octet-stream", "x-base-etag": base.eTag, "x-base-ctag": base.cTag, ...(force ? { "x-force": "1" } : {}) },
  })
  if (!response.ok) await failure(response, "The tracker could not be saved to OneDrive")
  return ((await response.json()) as { file: DriveFile }).file
}

export async function listDriveFiles(): Promise<DriveFile[]> {
  const response = await fetch("/api/onedrive/files", { cache: "no-store" })
  if (!response.ok) await failure(response, "OneDrive files could not be listed")
  return ((await response.json()) as { files: DriveFile[] }).files
}

export async function chooseDriveFile(id: string): Promise<DriveFile> {
  const response = await fetch("/api/onedrive/choose", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ id }) })
  if (!response.ok) await failure(response, "That file could not be chosen")
  return ((await response.json()) as { file: DriveFile }).file
}

export async function createDriveTracker(bytes: Uint8Array, name: string): Promise<DriveFile> {
  const response = await fetch(`/api/onedrive/files?name=${encodeURIComponent(name)}`, { method: "POST", headers: { "content-type": "application/octet-stream" }, body: new Blob([bytes.slice().buffer as ArrayBuffer]) })
  if (!response.ok) await failure(response, "The workbook could not be put into OneDrive")
  return ((await response.json()) as { file: DriveFile }).file
}

export async function disconnectDrive(): Promise<void> {
  await fetch("/api/onedrive/disconnect", { method: "POST" })
}
