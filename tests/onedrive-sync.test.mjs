import assert from "node:assert/strict"
import test from "node:test"
import { openDecision, pollDecision } from "../lib/onedrive-sync.ts"

const file = (cTag, id = "item-1") => ({ id, name: "Tracker.xlsx", eTag: `e-${cTag}`, cTag, modified: "2026-10-24T10:00:00Z", modifiedBy: "Someone", size: 1 })
const base = (cTag, savedRevision = 5) => ({ itemId: "item-1", eTag: `e-${cTag}`, cTag, savedRevision })

test("opening: same OneDrive version on this computer is used as is, even with unsaved edits", () => {
  assert.equal(openDecision({ revision: 7, exportedRevision: 1, drive: base("v1", 5) }, file("v1")), "use-local")
})

test("opening: a newer OneDrive version is loaded when nothing on this computer would be lost", () => {
  assert.equal(openDecision(null, file("v2")), "load-remote")
  assert.equal(openDecision({ revision: 5, exportedRevision: 1, drive: base("v1", 5) }, file("v2")), "load-remote")
  // A workbook opened from this computer and already downloaded is not lost either.
  assert.equal(openDecision({ revision: 3, exportedRevision: 3 }, file("v2")), "load-remote")
})

test("opening: a newer OneDrive version plus unsaved edits here asks the user", () => {
  assert.equal(openDecision({ revision: 6, exportedRevision: 1, drive: base("v1", 5) }, file("v2")), "ask")
  assert.equal(openDecision({ revision: 4, exportedRevision: 3 }, file("v2")), "ask")
})

test("while working: unchanged content is ignored, others' changes reload, clashes are flagged", () => {
  assert.equal(pollDecision(base("v1", 5), 9, { ...file("v1"), eTag: "renamed" }), "same")
  assert.equal(pollDecision(base("v1", 5), 5, file("v2")), "reload")
  assert.equal(pollDecision(base("v1", 5), 6, file("v2")), "conflict")
  assert.equal(pollDecision(base("v1", 5), 6, file("v1", "item-2")), "reload")
})
