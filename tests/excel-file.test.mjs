import assert from "node:assert/strict"
import test from "node:test"
import { changedOutside, fileDecision, writeExcelFile } from "../lib/excel-file.ts"

function mockHandle(fail=false) {
  let file=new File(["original"],"tracker.xlsx",{lastModified:10}),pending
  let writes=0,aborts=0
  return {getFile:async()=>file,replaceExternally:()=>{file=new File(["external edit"],"tracker.xlsx",{lastModified:30})},stats:()=>({writes,aborts}),createWritable:async()=>({write:async(blob)=>{writes++;if(fail)throw new Error("Disk write failed");pending=await blob.arrayBuffer()},close:async()=>{file=new File([pending],"tracker.xlsx",{lastModified:20})},abort:async()=>{aborts++}})}
}
test("Excel file save writes exact workbook bytes and returns the persisted file stamp",async()=>{
  const handle=mockHandle(),bytes=new Uint8Array([80,75,3,4,19])
  const stamp=await writeExcelFile(handle,bytes,{size:8,modified:10})
  assert.deepEqual(new Uint8Array(await (await handle.getFile()).arrayBuffer()),bytes)
  assert.deepEqual(stamp,{size:5,modified:20})
})
test("auto-save refuses to overwrite a workbook changed by another editor",async()=>{
  const handle=mockHandle();handle.replaceExternally()
  await assert.rejects(writeExcelFile(handle,new Uint8Array([1]),{size:8,modified:10}),/changed outside/)
  assert.equal(handle.stats().writes,0)
  assert.equal(await (await handle.getFile()).text(),"external edit")
})
test("a failed disk write aborts the writer and preserves the previous workbook",async()=>{
  const handle=mockHandle(true)
  await assert.rejects(writeExcelFile(handle,new Uint8Array([1]),{size:8,modified:10}),/Disk write failed/)
  assert.equal(handle.stats().aborts,1)
  assert.equal(await (await handle.getFile()).text(),"original")
})

test("reopening: an unchanged file keeps this computer's work; an edited file is loaded or, with unsaved work, asked about", () => {
  const known = { size: 100, modified: 1 }
  assert.equal(fileDecision(known, { size: 100, modified: 1 }, true), "same")
  assert.equal(fileDecision(known, { size: 120, modified: 2 }, false), "reload")
  assert.equal(fileDecision(known, { size: 100, modified: 2 }, true), "conflict")
  assert.equal(fileDecision(null, { size: 100, modified: 1 }, false), "reload")
})

test("a refused write because of outside changes is told apart from other write errors", () => {
  assert.equal(changedOutside(new Error("The Excel file changed outside the app. Open its latest version before saving again.")), true)
  assert.equal(changedOutside(new Error("The file is locked")), false)
})
