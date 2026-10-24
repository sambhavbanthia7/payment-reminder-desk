import type { WorkbookState } from "./debtors"
import type { DriveBase } from "./onedrive-sync"
import type { ExcelHandle, ExcelStamp } from "./excel-file"

/** `drive` is set when the workbook came from the shared OneDrive tracker. */
export type RecoveryDraft = { bytes: ArrayBuffer; state: WorkbookState; exportedRevision: number; savedAt: string; drive?: DriveBase }
const DATABASE = "jal-debtors-recovery-v1"
function open(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DATABASE, 2)
    request.onupgradeneeded = () => {
      const db = request.result
      if (!db.objectStoreNames.contains("drafts")) db.createObjectStore("drafts")
      if (!db.objectStoreNames.contains("files")) db.createObjectStore("files")
    }
    request.onsuccess = () => resolve(request.result)
    request.onerror = () => reject(request.error)
  })
}
export async function readDraft(): Promise<RecoveryDraft | null> {
  const db = await open()
  return new Promise((resolve, reject) => {
    const tx = db.transaction("drafts", "readonly")
    const request = tx.objectStore("drafts").get("current")
    request.onsuccess = () => resolve(request.result ?? null)
    request.onerror = () => reject(request.error)
    tx.oncomplete = () => db.close()
  })
}
export async function writeDraft(draft: RecoveryDraft): Promise<void> {
  const db = await open()
  return new Promise((resolve, reject) => {
    const tx = db.transaction("drafts", "readwrite")
    tx.objectStore("drafts").put(draft, "current")
    tx.oncomplete = () => { db.close(); resolve() }
    tx.onerror = () => { db.close(); reject(tx.error) }
    tx.onabort = () => { db.close(); reject(tx.error) }
  })
}

/** The tracker file this computer saves into, remembered between visits. Browsers keep file handles in IndexedDB. */
export type LinkedFile = { handle: ExcelHandle; name: string; stamp: ExcelStamp | null }

export async function readLinkedFile(): Promise<LinkedFile | null> {
  const db = await open()
  return new Promise((resolve, reject) => {
    const tx = db.transaction("files", "readonly")
    const request = tx.objectStore("files").get("tracker")
    request.onsuccess = () => resolve(request.result ?? null)
    request.onerror = () => reject(request.error)
    tx.oncomplete = () => db.close()
  })
}

export async function writeLinkedFile(file: LinkedFile | null): Promise<void> {
  const db = await open()
  return new Promise((resolve, reject) => {
    const tx = db.transaction("files", "readwrite")
    if (file) tx.objectStore("files").put(file, "tracker"); else tx.objectStore("files").delete("tracker")
    tx.oncomplete = () => { db.close(); resolve() }
    tx.onerror = () => { db.close(); reject(tx.error) }
  })
}
