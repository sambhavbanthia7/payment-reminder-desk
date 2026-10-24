import assert from "node:assert/strict"
import test from "node:test"
import { unzipSync, strFromU8 } from "fflate"

import { createLogWorkbook } from "../lib/log-export.ts"

test("Excel log export contains the sender, recipient, status, and numeric amount", () => {
  const bytes = createLogWorkbook([{
    id: "log-1",
    campaignId: "campaign-1",
    partyCode: "A-320",
    customerName: "A&B Traders",
    email: "billing@example.com",
    stage: "7-day window",
    invoiceNumbers: ["INV-1", "INV-2"],
    invoiceCount: 2,
    totalDue: 1234.5,
    status: "sent",
    sentAt: "2026-10-16T09:00:00.000Z",
    sentBy: "accounts@jalbath.com",
  }])
  const files = unzipSync(bytes)
  assert.ok(files["xl/workbook.xml"])
  const sheet = strFromU8(files["xl/worksheets/sheet1.xml"])
  assert.match(sheet, /accounts@jalbath\.com/)
  assert.match(sheet, /billing@example\.com/)
  assert.match(sheet, /A&amp;B Traders/)
  assert.match(sheet, /Accepted by Outlook/)
  assert.match(sheet, /<c r="I2" s="2"><v>1234\.5<\/v><\/c>/)
})
