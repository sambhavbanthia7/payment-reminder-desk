import { addActivity, appendReminderEntry, emailList, groupParties, localDate, money, partyKey, round, type Invoice, type Reminder, type WorkbookState } from "./debtors.ts"

export type EmailMode = "test" | "live"
/** What the dashboard knows about email sending. Never includes the password. */
export type EmailStatus = { configured: boolean; sender?: string; mode?: EmailMode; testRecipient?: string; sentCopy?: boolean; copyTo?: string[]; hourlyLimit?: number; problem?: string }
export type DraftPayload = Pick<Reminder, "id" | "type" | "party" | "email" | "subject" | "body" | "invoiceNumbers" | "amount">
export type SendResult = { id: string; status: "Accepted" | "Failed"; sentAt: string; sentTo: string; sentBy: string; mode: EmailMode; detail: string; repeat?: boolean }

/** Recipients and text for one draft. Test mode redirects every recipient to the test inbox. */
export function composeMail(draft: DraftPayload, mode: EmailMode, testRecipient: string): { recipients: string[]; subject: string; text: string } {
  const intended = emailList(draft.email)
  const recipients = mode === "test" ? emailList(testRecipient) : intended
  const subject = mode === "test" ? `[TEST] ${draft.subject}` : draft.subject
  const body = mode === "test" ? `TEST EMAIL. In live mode this would go to: ${intended.join("; ")}\n\n${draft.body}` : draft.body
  return { recipients, subject, text: body }
}

/**
 * Who the mail server delivers to: the visible recipients plus hidden copies (BCC). BCC addresses
 * never appear in the message headers. Test sends skip the copies so the team inbox stays clean.
 */
export function envelopeRecipients(recipients: string[], bcc: string[], mode: EmailMode): string[] {
  return [...new Set([...recipients, ...(mode === "live" ? bcc : [])].map((v) => v.toLowerCase()))]
}

export function validDraft(draft: DraftPayload): boolean {
  const list = emailList(draft?.email ?? "")
  return Boolean(draft?.id && draft.party && draft.subject?.trim() && draft.body?.trim() && list.length && list.every((v) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v)))
}

/** A draft is a snapshot. Refuse to send it when the balance or contact it was written for has changed. */
export function staleReason(draft: Reminder, invoices: Invoice[]): string {
  const rows = invoices.filter((i) => draft.invoiceIds.includes(i.id))
  if (rows.length !== draft.invoiceIds.length) return "An invoice in this draft is no longer in the workbook"
  if (Math.abs(round(rows.reduce((s, i) => s + i.outstanding, 0)) - draft.amount) > .01) return `Balance changed since the draft was prepared (draft ${money(draft.amount)})`
  const party = groupParties(invoices.filter((i) => i.outstanding > 0)).find((p) => p.key === partyKey(draft.party))
  if (!party || party.total <= 0) return "This customer no longer has an open balance"
  if (party.email !== draft.email) return "The customer's contact changed since the draft was prepared"
  return ""
}

/** Record provider results on the drafts and in the Activity Log. */
export function applySendResults(state: WorkbookState, results: SendResult[]): WorkbookState {
  if (!results.length) return state
  const byId = new Map(results.map((r) => [r.id, r]))
  const events = results.flatMap((r) => {
    const draft = state.reminders.find((d) => d.id === r.id)
    if (!draft) return []
    // A repeat means the server already handled this draft earlier and sent nothing now.
    const label = r.repeat ? "Send skipped" : r.status === "Accepted" ? (r.mode === "test" ? "Test email sent" : "Email sent") : "Email failed"
    return [{ ...addActivity(state, label, null, `${draft.type} · ${draft.invoiceNumbers.join(", ")} · ${money(draft.amount)} → ${r.sentTo} · by ${r.sentBy}${r.detail ? ` · ${r.detail}` : ""}`), party: draft.party, before: "Prepared", after: r.status }]
  })
  // A live reminder the mail server accepted is also written into "Last Reminder Sent" (column Q) of its invoices.
  const sentOn = new Map<string, string>()
  for (const r of results) {
    const draft = state.reminders.find((d) => d.id === r.id)
    if (!draft || r.mode !== "live" || r.status !== "Accepted" || r.repeat || draft.type !== "Reminder") continue
    for (const id of draft.invoiceIds) sentOn.set(id, localDate(r.sentAt))
  }
  return {
    ...state, revision: state.revision + 1, activity: [...state.activity, ...events],
    invoices: sentOn.size ? state.invoices.map((i) => sentOn.has(i.id) ? { ...i, legacyReminder: appendReminderEntry(i.legacyReminder, sentOn.get(i.id)!) } : i) : state.invoices,
    reminders: state.reminders.map((d) => {
      const r = byId.get(d.id)
      if (!r || (r.repeat && r.status === "Accepted" && r.mode === "test")) return d
      // A test send only proves the pipeline. It must not count as a customer reminder, so the draft stays Prepared.
      if (r.mode === "test") return { ...d, result: `Test ${r.status === "Accepted" ? "copy sent" : "failed"} to ${r.sentTo} by ${r.sentBy} · ${r.detail}` }
      return { ...d, status: r.status, sentAt: r.sentAt, sentTo: r.sentTo, sentBy: r.sentBy, result: r.detail }
    }),
  }
}
