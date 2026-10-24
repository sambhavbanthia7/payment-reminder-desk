export type ExcelStamp = { size:number; modified:number }
export type ExcelHandle = {
  name?:string
  getFile:()=>Promise<File>
  queryPermission?:(options:{mode:"read"|"readwrite"})=>Promise<"granted"|"denied"|"prompt">
  requestPermission?:(options:{mode:"read"|"readwrite"})=>Promise<"granted"|"denied"|"prompt">
  createWritable:()=>Promise<{write:(data:Blob)=>Promise<void>;close:()=>Promise<void>;abort?:()=>Promise<void>}>
}

export async function writeExcelFile(handle:ExcelHandle,bytes:Uint8Array,expected:ExcelStamp|null):Promise<ExcelStamp> {
  // Check immediately before writing, so another editor's file is not overwritten.
  const current=await handle.getFile()
  if(expected&&(current.size!==expected.size||current.lastModified!==expected.modified))throw new Error("The Excel file changed outside the app. Open its latest version before saving again.")
  const writer=await handle.createWritable()
  try {
    await writer.write(new Blob([bytes.slice().buffer as ArrayBuffer],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}))
    await writer.close()
  } catch(error) {
    await writer.abort?.().catch(()=>{})
    throw error
  }
  const saved=await handle.getFile()
  return {size:saved.size,modified:saved.lastModified}
}

export const sameStamp=(a:ExcelStamp|null,b:ExcelStamp|null)=>Boolean(a&&b&&a.size===b.size&&a.modified===b.modified)

/**
 * The tracker file on disk compared with what the app last wrote or read:
 * "same" - nothing changed outside the app; "reload" - changed (e.g. edited in Excel) and the app has
 * nothing unsaved, so take the file's version; "conflict" - changed and the app has unsaved changes too.
 */
export function fileDecision(known:ExcelStamp|null,current:ExcelStamp,unsaved:boolean):"same"|"reload"|"conflict" {
  if(sameStamp(known,current))return "same"
  return unsaved?"conflict":"reload"
}

/** A write refused because the file changed outside the app (as opposed to being temporarily blocked). */
export const changedOutside=(error:unknown)=>error instanceof Error&&error.message.startsWith("The Excel file changed outside the app")
