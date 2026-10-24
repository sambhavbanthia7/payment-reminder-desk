import assert from "node:assert/strict"
import test from "node:test"
import { appendReminderEntry, calculateInvoices, defaultSettings, invoiceId, reminderQueue } from "../lib/debtors.ts"
import { applySendResults, composeMail, envelopeRecipients, staleReason, validDraft } from "../lib/email-dispatch.ts"

const id = invoiceId("A&B Traders", "INV-1")
function fixture() {
  return {version:1,fileName:"tracker.xlsx",importedAt:"2026-10-16T10:00:00Z",revision:0,
    invoices:[{id,row:2,party:"A&B Traders",number:"INV-1",area:"Delhi",rse:"Sales",email:"accounts@example.com",issued:"2026-08-01",due:"2026-09-01",amount:1000,legacyReminder:"",remarks:""}],
    receipts:[],settings:{...defaultSettings},activity:[],followups:{},confirmations:[],importNotes:[],
    reminders:[{id:"d1",at:"2026-10-16T10:00:00Z",date:"2026-10-16",type:"Reminder",party:"A&B Traders",email:"accounts@example.com",invoiceIds:[id],invoiceNumbers:["INV-1"],amount:1000,stage:"Management Escalation",subject:"Reminder",body:"Please pay",status:"Prepared"}]}
}

test("test mode sends only to the test inbox and names the intended customer", () => {
  const draft = { ...fixture().reminders[0], email: "a@customer.com; b@customer.com" }
  const test = composeMail(draft, "test", "me@jal.example")
  assert.deepEqual(test.recipients, ["me@jal.example"])
  assert.equal(test.subject, "[TEST] Reminder")
  assert.match(test.text, /a@customer\.com; b@customer\.com/)
  const live = composeMail(draft, "live", "me@jal.example")
  assert.deepEqual(live.recipients, ["a@customer.com", "b@customer.com"])
  assert.equal(live.subject, "Reminder")
  assert.equal(live.text, "Please pay")
})

test("drafts with invalid recipients are refused", () => {
  assert.equal(validDraft(fixture().reminders[0]), true)
  assert.equal(validDraft({ ...fixture().reminders[0], email: "not-an-email" }), false)
})

test("a draft is blocked when the balance or contact changed after preparation", () => {
  const state = fixture()
  assert.equal(staleReason(state.reminders[0], calculateInvoices(state, "2026-10-16")), "")
  state.receipts.push({id:"r1",row:0,date:"2026-10-16",party:"A&B Traders",invoice:"INV-1",amount:200,reference:"",remarks:""})
  assert.match(staleReason(state.reminders[0], calculateInvoices(state, "2026-10-16")), /Balance changed/)
  const moved = fixture(); moved.invoices[0].email = "new@example.com"
  assert.match(staleReason(moved.reminders[0], calculateInvoices(moved, "2026-10-16")), /contact changed/)
})

test("test sends leave the draft Prepared; live sends mark it Accepted and count as the last reminder", () => {
  const state = fixture()
  const tested = applySendResults(state, [{id:"d1",status:"Accepted",sentAt:"2026-10-16T11:00:00Z",sentTo:"me@jal.example",sentBy:"Priya",mode:"test",detail:"Accepted by the mail server"}])
  assert.equal(tested.reminders[0].status, "Prepared")
  assert.equal(tested.activity.at(-1).action, "Test email sent")
  const live = applySendResults(state, [{id:"d1",status:"Accepted",sentAt:"2026-10-16T11:00:00Z",sentTo:"accounts@example.com",sentBy:"Priya",mode:"live",detail:"Accepted by the mail server"}])
  assert.equal(live.reminders[0].status, "Accepted")
  assert.equal(live.reminders[0].sentTo, "accounts@example.com")
  assert.equal(live.reminders[0].sentBy, "Priya")
  assert.match(live.activity.at(-1).detail, /by Priya/)
  assert.equal(calculateInvoices(live, "2026-10-20")[0].lastReminder, "2026-10-16")
  assert.equal(reminderQueue(live, calculateInvoices(live, "2026-10-20"), "2026-10-20").length, 0)
  const repeat = applySendResults(tested, [{id:"d1",status:"Accepted",sentAt:"2026-10-16T11:00:00Z",sentTo:"me@jal.example",sentBy:"Priya",mode:"test",detail:"Already sent earlier; not sent again",repeat:true}])
  assert.equal(repeat.activity.at(-1).action, "Send skipped")
  assert.equal(repeat.reminders[0].result, tested.reminders[0].result)
  const failed = applySendResults(state, [{id:"d1",status:"Failed",sentAt:"2026-10-16T11:00:00Z",sentTo:"accounts@example.com",sentBy:"Priya",mode:"live",detail:"The mail server refused the recipient"}])
  assert.equal(failed.reminders[0].status, "Failed")
  assert.equal(failed.activity.at(-1).action, "Email failed")
})

test("live sends add the hidden team copy to the envelope only; test sends skip it", () => {
  assert.deepEqual(envelopeRecipients(["a@customer.com", "b@customer.com"], ["finance@jal.example"], "live"), ["a@customer.com", "b@customer.com", "finance@jal.example"])
  assert.deepEqual(envelopeRecipients(["me@jal.example"], ["finance@jal.example"], "test"), ["me@jal.example"])
  assert.deepEqual(envelopeRecipients(["Finance@JAL.example"], ["finance@jal.example"], "live"), ["finance@jal.example"])
})

test("test mode can send to several test inboxes", () => {
  const mail = composeMail(fixture().reminders[0], "test", "one@test.example, two@test.example")
  assert.deepEqual(mail.recipients, ["one@test.example", "two@test.example"])
})

test("a live reminder adds a dated MAIL entry to Last Reminder Sent, keeping what was there", () => {
  const state = fixture(); state.invoices[0].legacyReminder = "28.09.26, 06.10.26"
  const sent = { id: "d1", status: "Accepted", sentAt: "2026-10-24T06:00:00Z", sentTo: "accounts@example.com", sentBy: "Priya", mode: "live", detail: "Accepted by the mail server" }
  const live = applySendResults(state, [sent])
  assert.equal(live.invoices[0].legacyReminder, "28.09.26, 06.10.26, 24.10.26 MAIL")
  assert.equal(applySendResults(live, [{ ...sent, repeat: true }]).invoices[0].legacyReminder, "28.09.26, 06.10.26, 24.10.26 MAIL")
  assert.equal(applySendResults(state, [{ ...sent, mode: "test" }]).invoices[0].legacyReminder, "28.09.26, 06.10.26")
  assert.equal(applySendResults(state, [{ ...sent, status: "Failed" }]).invoices[0].legacyReminder, "28.09.26, 06.10.26")
  const confirmation = fixture(); confirmation.reminders[0].type = "Balance confirmation"
  assert.equal(applySendResults(confirmation, [sent]).invoices[0].legacyReminder, "")
})

test("the column Q entry is written once per day and starts an empty cell cleanly", () => {
  assert.equal(appendReminderEntry("", "2026-10-24"), "24.10.26 MAIL")
  assert.equal(appendReminderEntry("30.09.26 MAIL", "2026-10-24"), "30.09.26 MAIL, 24.10.26 MAIL")
  assert.equal(appendReminderEntry("24.10.26 MAIL", "2026-10-24"), "24.10.26 MAIL")
})
