export type InvoiceInput = {
  id: string; row: number; party: string; area: string; rse: string; email: string
  number: string; issued: string; due: string; amount: number; legacyReminder: string; remarks: string
}
export type Receipt = { id: string; row: number; date: string; party: string; invoice: string; amount: number; reference: string; remarks: string }
export type Followup = { flag: boolean; escalation: string; hold: boolean; note: string }
export type Activity = { id: string; at: string; action: string; invoiceId: string; party: string; invoice: string; detail: string; before: string; after: string }
export type Reminder = {
  id: string; at: string; date: string; type: "Reminder" | "Balance confirmation"; party: string; email: string
  invoiceIds: string[]; invoiceNumbers: string[]; amount: number; stage: string; subject: string; body: string
  status: "Prepared" | "Cancelled" | "Accepted" | "Failed"
  sentAt?: string; sentTo?: string; sentBy?: string; result?: string
}
export type Confirmation = { id: string; at: string; party: string; date: string; amount: number; status: "Confirmed" | "Disputed"; reference: string; note: string }
export type Settings = { reminderSubject: string; reminderBody: string; balanceSubject: string; balanceBody: string; preDueDays: number; repeatDays: number }
export type WorkbookState = {
  version: 1; fileName: string; importedAt: string; revision: number; invoices: InvoiceInput[]; receipts: Receipt[]
  activity: Activity[]; reminders: Reminder[]; followups: Record<string, Followup>; confirmations: Confirmation[]
  settings: Settings; importNotes: string[]
}
export type Invoice = InvoiceInput & { received: number; outstanding: number; overdue: number; age: number; daysToDue: number; stage: string; ageing: string; status: string; risk: string; issues: string[]; lastReminder: string; followup: Followup }
export type Party = { key: string; name: string; invoices: Invoice[]; total: number; overdue: number; maxOverdue: number; risk: string; emails: string[]; email: string; issues: string[] }
export const defaultSettings: Settings = {
  reminderSubject: "Payment reminder: {{customer_name}} · {{total_outstanding}}",
  reminderBody: "Dear {{customer_name}},\n\nPlease find the outstanding invoice details as of {{as_of_date}}:\n\n{{invoice_list}}\n\nTotal outstanding: {{total_outstanding}}\n\nPlease arrange payment and share the payment reference. If you have already paid, please send the details so we can reconcile our records.\n\nRegards,\nJAL Accounts Receivable",
  balanceSubject: "Balance confirmation as of {{as_of_date}} · {{customer_name}}",
  balanceBody: "Dear {{customer_name}},\n\nOur records show a balance of {{total_outstanding}} outstanding as of {{as_of_date}}.\n\n{{invoice_list}}\n\nPlease confirm this balance or share details of any differences and payments pending reconciliation.\n\nRegards,\nJAL Accounts Receivable",
  preDueDays: 5, repeatDays: 7,
}
export const money = (value: number) => new Intl.NumberFormat("en-IN", { style: "currency", currency: "INR", maximumFractionDigits: 2 }).format(value)
export function shortMoney(value: number): string {
  return Math.abs(value) >= 1e7 ? `₹${(value / 1e7).toFixed(2)} Cr` : Math.abs(value) >= 1e5 ? `₹${(value / 1e5).toFixed(2)} L` : money(value)
}
export const round = (value: number) => Math.round((value + Number.EPSILON) * 100) / 100
export const partyKey = (name: string) => name.toLowerCase()
export const allocationKey = (party: string, number: string) => JSON.stringify([partyKey(party), number.toLowerCase()])
export const invoiceId = (party: string, number: string) => encodeURIComponent(allocationKey(party, number))
export function today(): string { return localDate(new Date().toISOString()) }
export function localDate(iso: string): string { const d = new Date(iso); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}` }
export function dayNumber(value: string): number { return Date.parse(`${value}T00:00:00Z`) / 86400000 }
export function validDate(value: string): boolean { return /^\d{4}-\d{2}-\d{2}$/.test(value) && Number.isFinite(dayNumber(value)) && new Date(dayNumber(value) * 86400000).toISOString().slice(0, 10) === value }
export function displayDate(value: string): string { return validDate(value) ? new Date(`${value}T00:00:00Z`).toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric", timeZone: "UTC" }) : "Date missing" }
export function parseLegacyDates(value: string): string[] {
  const dates = [...value.matchAll(/\b(\d{1,2})[./-](\d{1,2})[./-](\d{2}|\d{4})\b/g)].map((m) => `${m[3].length === 2 ? `20${m[3]}` : m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`)
  if (validDate(value)) dates.push(value)
  return dates.filter(validDate).sort()
}
/**
 * Add an emailed reminder to the tracker's "Last Reminder Sent" text (column Q), in the team's own
 * style: "28.09.26, 06.10.26" plus "24.10.26 MAIL". Existing text is kept, and the same entry is never added twice.
 */
export function appendReminderEntry(text: string, date: string): string {
  if (!validDate(date)) return text
  const entry = `${date.slice(8, 10)}.${date.slice(5, 7)}.${date.slice(2, 4)} MAIL`
  const current = text.trim()
  if (current.includes(entry)) return text
  return current ? `${current}, ${entry}` : entry
}
export function stageFor(overdue: number): string { return overdue < 5 ? "Pre-Due / Due" : overdue < 15 ? "1st Reminder" : overdue < 30 ? "2nd Reminder" : overdue < 45 ? "3rd Reminder / CFO" : "Management Escalation" }
export function riskFor(overdue: number, flag = false, escalation = ""): string { return overdue >= 45 || /management|legal/i.test(escalation) ? "Critical" : overdue >= 30 || flag || /cfo/i.test(escalation) ? "High" : overdue >= 15 ? "Watch" : "Standard" }
export const riskRank: Record<string, number> = { Critical: 4, High: 3, Watch: 2, Standard: 1, Closed: 0 }
export const ageingLabels = ["Not due", "1–30 days", "31–45 days", "46–60 days", "61–90 days", "90+ days"]
export function ageingFor(days: number): string { return days <= 0 ? ageingLabels[0] : days <= 30 ? ageingLabels[1] : days <= 45 ? ageingLabels[2] : days <= 60 ? ageingLabels[3] : days <= 90 ? ageingLabels[4] : ageingLabels[5] }
export function emailList(value: string): string[] { return [...new Set(value.split(/[;,\s]+/).map((v) => v.trim().toLowerCase()).filter(Boolean))] }
export function validEmail(value: string): boolean { const list = emailList(value); return list.length > 0 && list.every((v) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v)) }

export function calculateInvoices(state: WorkbookState, asOf: string): Invoice[] {
  const amounts = new Map<string, number>()
  for (const r of state.receipts) { const key = allocationKey(r.party, r.invoice); amounts.set(key, round((amounts.get(key) ?? 0) + r.amount)) }
  const duplicates = new Map<string, number>()
  for (const i of state.invoices) { const key = allocationKey(i.party, i.number); duplicates.set(key, (duplicates.get(key) ?? 0) + 1) }
  const latestSent = new Map<string, string>()
  for (const r of state.reminders.filter((r) => r.type === "Reminder" && r.status === "Accepted")) for (const id of r.invoiceIds) { const sent = r.sentAt ? localDate(r.sentAt) : r.date; if (sent > (latestSent.get(id) ?? "")) latestSent.set(id, sent) }
  return state.invoices.map((i) => {
    const received = round(amounts.get(allocationKey(i.party, i.number)) ?? 0)
    const outstanding = round(Math.max(0, i.amount - received))
    const datesValid = validDate(i.due) && validDate(i.issued)
    const daysToDue = datesValid ? dayNumber(i.due) - dayNumber(asOf) : NaN
    const overdue = outstanding > 0 && datesValid ? Math.max(0, -daysToDue) : 0
    const followup = state.followups[i.id] ?? { flag: false, escalation: "", hold: false, note: "" }
    const issues = [
      ...(!datesValid ? ["Missing or invalid date"] : []),
      ...(!validEmail(i.email) ? ["Missing or invalid email"] : []),
      ...((duplicates.get(allocationKey(i.party, i.number)) ?? 0) > 1 ? ["Duplicate party + invoice"] : []),
      ...(received > Math.max(0, i.amount) + .01 ? ["Receipt exceeds invoice"] : []),
      ...(i.amount < 0 ? ["Credit entry"] : []),
    ]
    const legacyDate = parseLegacyDates(i.legacyReminder).filter((d) => d <= asOf).at(-1) ?? ""
    return { ...i, received, outstanding, overdue, daysToDue, age: datesValid ? Math.max(0, dayNumber(asOf) - dayNumber(i.issued)) : 0,
      stage: outstanding <= 0 ? "Closed" : !datesValid ? "Pending" : stageFor(overdue), ageing: outstanding <= 0 ? "Paid" : !datesValid ? "Unknown" : ageingFor(overdue),
      status: i.amount < 0 ? "Credit entry" : outstanding <= 0 ? "Paid" : received > 0 ? "Part paid" : overdue > 0 ? "Overdue" : "Open",
      risk: outstanding <= 0 ? "Closed" : riskFor(overdue, followup.flag, followup.escalation), issues,
      lastReminder: [legacyDate, latestSent.get(i.id) ?? ""].sort().at(-1) ?? "", followup }
  })
}
export function groupParties(invoices: Invoice[]): Party[] {
  const groups = new Map<string, Invoice[]>()
  for (const invoice of invoices) { const key = partyKey(invoice.party); groups.set(key, [...(groups.get(key) ?? []), invoice]) }
  return [...groups].map(([key, rows]) => {
    const emails = [...new Set(rows.map((i) => i.email.trim().toLowerCase()).filter(Boolean))]
    const open = rows.filter((i) => i.outstanding > 0)
    const activeEmails = [...new Set(open.map((i) => i.email.trim().toLowerCase()).filter(Boolean))]
    const usable = activeEmails.length ? activeEmails : emails
    const issues = [...new Set(rows.flatMap((i) => i.issues.filter((v) => v !== "Missing or invalid email")))]
    if (usable.length > 1) issues.push("Conflicting contacts")
    if (usable.length !== 1 || !validEmail(usable[0])) issues.push("Contact needs review")
    return { key, name: rows[0].party, invoices: rows, total: round(rows.reduce((s, i) => s + i.outstanding, 0)), overdue: round(rows.filter((i) => i.overdue > 0).reduce((s, i) => s + i.outstanding, 0)), maxOverdue: Math.max(0, ...open.map((i) => i.overdue)), risk: [...rows].sort((a,b) => riskRank[b.risk] - riskRank[a.risk])[0].risk, emails: usable, email: usable.length === 1 && validEmail(usable[0]) ? usable[0] : "", issues }
  }).sort((a,b) => riskRank[b.risk] - riskRank[a.risk] || b.total - a.total)
}
export function reminderQueue(state: WorkbookState, invoices: Invoice[], date: string): Party[] {
  const eligible = invoices.filter((i) => {
    if (i.outstanding <= 0 || i.followup.hold || !Number.isFinite(i.daysToDue) || i.daysToDue > state.settings.preDueDays) return false
    const prepared = state.reminders.some((r) => r.type === "Reminder" && r.date === date && r.status !== "Cancelled" && r.status !== "Failed" && r.invoiceIds.includes(i.id))
    const sinceLast = i.lastReminder ? dayNumber(date) - dayNumber(i.lastReminder) : Infinity
    return !prepared && sinceLast >= state.settings.repeatDays
  })
  const contacts=new Map(groupParties(calculateInvoices(state,date).filter((i)=>i.outstanding>0)).map((p)=>[p.key,p]))
  return groupParties(eligible).map((p)=>{
    const full=contacts.get(p.key)
    return full?{...p,email:full.email,emails:full.emails,issues:[...new Set([...p.issues,...full.issues.filter((issue)=>issue==="Conflicting contacts"||issue==="Contact needs review")])]}:p
  })
}
export function renderMessage(settings: Settings, party: Party, date: string, type: Reminder["type"]): { subject: string; body: string } {
  const open = party.invoices.filter((i) => i.outstanding > 0)
  const values: Record<string, string> = {
    customer_name: party.name.trim(), total_outstanding: money(party.total), total_due: money(party.total), as_of_date: displayDate(date),
    invoice_count: String(open.length), email: party.email, stage: [...open].sort((a,b) => b.overdue - a.overdue)[0]?.stage ?? "Balance confirmation",
    invoice_list: open.map((i) => `${i.number} | Invoice: ${displayDate(i.issued)} | Due: ${displayDate(i.due)} | Outstanding: ${money(i.outstanding)}`).join("\n"),
  }
  const replace = (text: string) => text.replace(/\{\{\s*(\w+)\s*\}\}/g, (match,key:string) => values[key] ?? match)
  return { subject: replace(type === "Reminder" ? settings.reminderSubject : settings.balanceSubject), body: replace(type === "Reminder" ? settings.reminderBody : settings.balanceBody) }
}
export function addActivity(state: WorkbookState, action: string, invoice: Pick<InvoiceInput,"id"|"party"|"number"> | null, detail: string, before = "", after = ""): Activity {
  return { id: crypto.randomUUID(), at: new Date().toISOString(), action, invoiceId: invoice?.id ?? "", party: invoice?.party ?? "", invoice: invoice?.number ?? "", detail, before, after }
}
export function recordReceipt(state: WorkbookState, invoice: Invoice, input: { date: string; amount: number; reference: string; remarks: string }, receiptId?: string): WorkbookState {
  if (!validDate(input.date)) throw new Error("Choose a valid payment date")
  if (!Number.isFinite(input.amount) || input.amount <= 0) throw new Error("Payment amount must be greater than zero")
  if (invoice.amount <= 0 || invoice.issues.includes("Duplicate party + invoice")) throw new Error("Resolve this invoice's credit or duplicate entry before allocating payments")
  const previous = receiptId ? state.receipts.find((r) => r.id === receiptId) : undefined
  if (receiptId && (!previous || allocationKey(previous.party, previous.invoice) !== allocationKey(invoice.party, invoice.number))) throw new Error("The receipt does not belong to this invoice")
  const available = round(invoice.amount - invoice.received + (previous?.amount ?? 0))
  if (input.amount > available + .01) throw new Error(`Payment exceeds the available invoice balance of ${money(available)}`)
  const receipt: Receipt = { id: previous?.id ?? crypto.randomUUID(), row: previous?.row ?? 0, party: invoice.party, invoice: invoice.number, ...input, amount: round(input.amount) }
  const activity = addActivity(state, previous ? "Payment updated" : "Payment recorded", invoice, `${money(receipt.amount)} · ${receipt.reference || "No reference"}`, previous ? JSON.stringify(previous) : "", JSON.stringify(receipt))
  return { ...state, revision: state.revision + 1, receipts: previous ? state.receipts.map((r) => r.id === previous.id ? receipt : r) : [...state.receipts, receipt], activity: [...state.activity, activity] }
}
export function updateContact(state: WorkbookState, party: string, email: string): WorkbookState {
  if (!validEmail(email)) throw new Error("Enter valid email addresses, separated by semicolons")
  const nextEmail = emailList(email).join("; ")
  const affected = state.invoices.filter((i) => partyKey(i.party) === partyKey(party))
  const event = addActivity(state, "Contact updated", affected[0] ?? null, `Updated contact on ${affected.length} invoice rows`, [...new Set(affected.map((i) => i.email))].join("; "), nextEmail)
  return { ...state, revision: state.revision + 1, invoices: state.invoices.map((i) => partyKey(i.party) === partyKey(party) ? { ...i, email: nextEmail } : i), activity: [...state.activity, event] }
}
