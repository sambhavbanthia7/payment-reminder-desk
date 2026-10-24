import assert from "node:assert/strict"
import test from "node:test"
import { allocationKey, calculateInvoices, defaultSettings, groupParties, invoiceId, parseLegacyDates, recordReceipt, reminderQueue, renderMessage, stageFor, updateContact } from "../lib/debtors.ts"

function fixture() {
  return {version:1,fileName:"tracker.xlsx",importedAt:"2026-10-16T10:00:00Z",revision:0,
    invoices:[{id:invoiceId("A&B Traders","INV-1"),row:2,party:"A&B Traders",number:"INV-1",area:"Delhi",rse:"Sales",email:"accounts@example.com",issued:"2026-08-01",due:"2026-09-01",amount:1000,legacyReminder:"",remarks:""},
      {id:invoiceId("Other Party","INV-1"),row:3,party:"Other Party",number:"INV-1",area:"Mumbai",rse:"Sales",email:"other@example.com",issued:"2026-10-01",due:"2026-10-20",amount:800,legacyReminder:"",remarks:""}],
    receipts:[{id:"r1",row:2,date:"2026-10-01",party:"A&B Traders",invoice:"INV-1",amount:200,reference:"UTR1",remarks:""}],
    settings:{...defaultSettings},activity:[],reminders:[],followups:{},confirmations:[],importNotes:[]}
}
test("receipt allocation matches both party and invoice, retaining Excel's space distinction",()=>{
  const state=fixture();state.receipts.push({...state.receipts[0],id:"unmatched",party:"A&B Traders ",amount:100})
  const rows=calculateInvoices(state,"2026-10-16")
  assert.equal(rows[0].received,200);assert.equal(rows[0].outstanding,800);assert.equal(rows[1].received,0)
  assert.equal(allocationKey("PARTY","inv"),allocationKey("party","INV"))
})
test("SOP stages handle exact 5, 15, 30 and 45 day boundaries",()=>{
  assert.deepEqual([4,5,14,15,29,30,44,45].map(stageFor),["Pre-Due / Due","1st Reminder","1st Reminder","2nd Reminder","2nd Reminder","3rd Reminder / CFO","3rd Reminder / CFO","Management Escalation"])
  const rows=calculateInvoices(fixture(),"2026-10-16");assert.equal(rows[0].risk,"Critical");assert.equal(rows[1].overdue,0)
})
test("part payments and receipt corrections recalculate balance with immutable audit entries",()=>{
  const initial=fixture(),invoice=calculateInvoices(initial,"2026-10-16")[0]
  const paid=recordReceipt(initial,invoice,{date:"2026-10-16",amount:300,reference:"UTR2",remarks:"Part payment"})
  assert.equal(calculateInvoices(paid,"2026-10-16")[0].outstanding,500);assert.equal(initial.receipts.length,1)
  const corrected=recordReceipt(paid,calculateInvoices(paid,"2026-10-16")[0],{date:"2026-10-16",amount:250,reference:"UTR2 corrected",remarks:"Correction"},paid.receipts[1].id)
  assert.equal(calculateInvoices(corrected,"2026-10-16")[0].outstanding,550);assert.equal(corrected.activity.length,2);assert.match(corrected.activity[1].before,/300/)
  assert.throws(()=>recordReceipt(corrected,calculateInvoices(corrected,"2026-10-16")[0],{date:"2026-10-16",amount:551,reference:"",remarks:""}),/exceeds/)
})
test("paid and credit rows stay out of reminders; duplicate invoice pairs block payment entry",()=>{
  const state=fixture();state.receipts[0].amount=1000;state.invoices[1].amount=-100
  assert.equal(reminderQueue(state,calculateInvoices(state,"2026-10-16"),"2026-10-16").length,0)
  const duplicate=fixture();duplicate.invoices.push({...duplicate.invoices[0],id:"dup",row:4})
  const row=calculateInvoices(duplicate,"2026-10-16")[0];assert.ok(row.issues.includes("Duplicate party + invoice"));assert.throws(()=>recordReceipt(duplicate,row,{date:"2026-10-16",amount:10,reference:"",remarks:""}),/duplicate/)
})
test("preparing a reminder suppresses same-day duplicates without marking it sent",()=>{
  const state=fixture(),date="2026-10-16",rows=calculateInvoices(state,date)
  assert.equal(reminderQueue(state,rows,date).length,2)
  state.reminders.push({id:"draft",at:date,date,type:"Reminder",party:rows[0].party,email:rows[0].email,invoiceIds:[rows[0].id],invoiceNumbers:["INV-1"],amount:800,stage:"Management Escalation",subject:"Test",body:"Test",status:"Prepared"})
  assert.equal(reminderQueue(state,calculateInvoices(state,date),date).length,1)
  assert.equal(calculateInvoices(state,date)[0].lastReminder,"")
  state.reminders[0].status="Cancelled";assert.equal(reminderQueue(state,calculateInvoices(state,date),date).length,2)
})
test("legacy dates and held invoices respect reminder cadence",()=>{
  const state=fixture(),date="2026-10-16";state.invoices[0].legacyReminder="28.09.26, 14.10.26 MAIL"
  state.followups[state.invoices[1].id]={flag:false,escalation:"",hold:true,note:"Dispute"}
  assert.deepEqual(parseLegacyDates(state.invoices[0].legacyReminder),["2026-09-28","2026-10-14"])
  assert.equal(reminderQueue(state,calculateInvoices(state,date),date).length,0)
})
test("contact changes apply to all party rows; templates contain current invoices and balance",()=>{
  const state=fixture();state.invoices.push({...state.invoices[0],id:invoiceId("A&B Traders","INV-2"),number:"INV-2",row:4,amount:500})
  const changed=updateContact(state,"a&b traders","New@example.com; owner@example.com")
  assert.equal(changed.invoices[0].email,"new@example.com; owner@example.com");assert.equal(changed.invoices[2].email,changed.invoices[0].email)
  const party=groupParties(calculateInvoices(changed,"2026-10-16")).find((p)=>p.name==="A&B Traders")
  const message=renderMessage(changed.settings,party,"2026-10-16","Balance confirmation")
  assert.match(message.body,/INV-1/);assert.match(message.body,/INV-2/);assert.match(message.body,/1,300/);assert.doesNotMatch(message.body,/\{\{/)
})
