export type SheetTable = {
  fileName: string
  sheetName: string
  headers: string[]
  rows: Record<string, unknown>[]
}

export type DuesMapping = {
  accountName: string
  accountCode: string
  creditDays: string
  invoiceDate: string
  invoiceNumber: string
  transactionType: string
  balance: string
  ageDays: string
}

export type MasterMapping = {
  partyCode: string
  name: string
  email: string
}

export type CampaignConfig = {
  firstThreshold: number
  secondThreshold: number
  defaultCreditDays: number
  subject: string
  body: string
  asOfDate: string
}

export type ReminderInvoice = {
  id: string
  sourceRowNumber: number
  sourceRow: Record<string, unknown>
  partyCode: string
  customerName: string
  email: string | null
  invoiceNumber: string
  invoiceDate: Date
  dueDate: Date
  daysRemaining: number
  balance: number
  creditDays: number
  usedDefaultCredit: boolean
  issue: "unmatched" | "missing_email" | "invalid_email" | "negative_credit" | null
}

export type CustomerReminder = {
  partyCode: string
  customerName: string
  email: string | null
  invoices: ReminderInvoice[]
  totalDue: number
  nearestDue: number
  issue: ReminderInvoice["issue"]
  usedDefaultCredit: boolean
}

export type SendLog = {
  id: string
  partyCode: string
  customerName: string
  email: string
  stage: string
  invoiceNumbers: string[]
  invoiceCount: number
  totalDue: number
  status: "sent" | "failed" | "previewed"
  sentAt: string
  errorMessage?: string
}

const DAY_MS = 86_400_000

export function parseDate(value: unknown): Date | null {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return startOfDay(value)
  if (typeof value === "number" && Number.isFinite(value)) {
    return startOfDay(new Date(Date.UTC(1899, 11, 30) + value * DAY_MS))
  }
  const raw = String(value ?? "").trim()
  if (!raw) return null
  const dmy = raw.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2}|\d{4})$/)
  if (dmy) {
    const year = Number(dmy[3]) < 100 ? 2000 + Number(dmy[3]) : Number(dmy[3])
    const parsed = new Date(Date.UTC(year, Number(dmy[2]) - 1, Number(dmy[1])))
    return Number.isNaN(parsed.getTime()) ? null : parsed
  }
  const iso = new Date(raw)
  return Number.isNaN(iso.getTime()) ? null : startOfDay(iso)
}

export function parseAmount(value: unknown): number {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0
  const cleaned = String(value ?? "")
    .replace(/<[^>]*>/g, "")
    .replace(/[^0-9.()-]/g, "")
    .replace(/^\((.*)\)$/, "-$1")
  const parsed = Number(cleaned)
  return Number.isFinite(parsed) ? parsed : 0
}

export function deriveReportDate(table: SheetTable, mapping: DuesMapping): Date {
  const candidates = table.rows
    .map((row) => {
      const date = parseDate(row[mapping.invoiceDate])
      const age = Number(row[mapping.ageDays])
      if (!date || !Number.isFinite(age)) return null
      return addDays(date, age)
    })
    .filter((date): date is Date => Boolean(date))
  if (!candidates.length) return startOfDay(new Date())
  const counts = new Map<number, number>()
  for (const date of candidates) counts.set(date.getTime(), (counts.get(date.getTime()) ?? 0) + 1)
  const winner = [...counts.entries()].sort((a, b) => b[1] - a[1])[0]
  return new Date(winner[0])
}

export function buildInvoices(
  dues: SheetTable,
  master: SheetTable,
  duesMap: DuesMapping,
  masterMap: MasterMapping,
  config: CampaignConfig,
): ReminderInvoice[] {
  const asOf = parseDate(config.asOfDate) ?? deriveReportDate(dues, duesMap)
  const masterByCode = new Map(
    master.rows
      .map((row) => [normaliseCode(row[masterMap.partyCode]), row] as const)
      .filter(([code]) => Boolean(code)),
  )

  return dues.rows.flatMap((row, index) => {
    const transactionType = String(row[duesMap.transactionType] ?? "").trim().toUpperCase()
    if (transactionType !== "GST INVOICE") return []
    const balance = parseAmount(row[duesMap.balance])
    if (balance <= 0) return []
    const invoiceDate = parseDate(row[duesMap.invoiceDate])
    if (!invoiceDate) return []
    const partyCode = normaliseCode(row[duesMap.accountCode])
    const rawCredit = String(row[duesMap.creditDays] ?? "").trim()
    const usedDefaultCredit = rawCredit === ""
    const creditDays = usedDefaultCredit ? config.defaultCreditDays : Number(rawCredit)
    const negativeCredit = !Number.isFinite(creditDays) || creditDays < 0
    const safeCredit = negativeCredit ? 0 : creditDays
    const dueDate = addDays(invoiceDate, safeCredit)
    const masterRow = masterByCode.get(partyCode)
    const email = masterRow ? cleanEmail(masterRow[masterMap.email]) : null
    const fallbackName = stripLocation(String(row[duesMap.accountName] ?? partyCode))
    const customerName = String(masterRow?.[masterMap.name] ?? fallbackName).trim() || partyCode
    const issue: ReminderInvoice["issue"] = negativeCredit
      ? "negative_credit"
      : !masterRow
        ? "unmatched"
        : !email
          ? "missing_email"
          : !isValidEmail(email)
            ? "invalid_email"
            : null

    return [{
      id: `${partyCode}-${String(row[duesMap.invoiceNumber] ?? index)}`,
      sourceRowNumber: index + 2,
      sourceRow: row,
      partyCode,
      customerName,
      email,
      invoiceNumber: String(row[duesMap.invoiceNumber] ?? `Row ${index + 1}`).trim(),
      invoiceDate,
      dueDate,
      daysRemaining: differenceInDays(dueDate, asOf),
      balance,
      creditDays: safeCredit,
      usedDefaultCredit,
      issue,
    }]
  })
}

export function groupCustomers(invoices: ReminderInvoice[]): CustomerReminder[] {
  const grouped = new Map<string, ReminderInvoice[]>()
  for (const invoice of invoices) {
    const list = grouped.get(invoice.partyCode) ?? []
    list.push(invoice)
    grouped.set(invoice.partyCode, list)
  }
  return [...grouped.entries()].map(([partyCode, rows]) => ({
    partyCode,
    customerName: rows[0].customerName,
    email: rows[0].email,
    invoices: rows.sort((a, b) => a.dueDate.getTime() - b.dueDate.getTime()),
    totalDue: rows.reduce((sum, row) => sum + row.balance, 0),
    nearestDue: Math.min(...rows.map((row) => row.daysRemaining)),
    issue: rows.find((row) => row.issue)?.issue ?? null,
    usedDefaultCredit: rows.some((row) => row.usedDefaultCredit),
  }))
}

export function applyEmailOverrides(
  invoices: ReminderInvoice[],
  overrides: Record<string, string>,
): ReminderInvoice[] {
  return invoices.map((invoice) => {
    const override = overrides[invoice.partyCode]?.trim().toLowerCase()
    if (!override || !isValidEmail(override) || invoice.issue === "negative_credit") return invoice
    return {
      ...invoice,
      email: override,
      issue: invoice.issue === "unmatched" || invoice.issue === "missing_email" || invoice.issue === "invalid_email"
        ? null
        : invoice.issue,
    }
  })
}

export function detectMapping(headers: string[], type: "dues"): DuesMapping
export function detectMapping(headers: string[], type: "master"): MasterMapping
export function detectMapping(headers: string[], type: "dues" | "master"): DuesMapping | MasterMapping {
  const find = (...aliases: string[]) => {
    const normalised = aliases.map(normaliseHeader)
    return headers.find((header) => normalised.includes(normaliseHeader(header))) ?? ""
  }
  if (type === "dues") {
    return {
      accountName: find("Account Name", "Party Name", "Customer Name", "Name"),
      accountCode: find("A/C Code", "Account Code", "Party Code", "Customer Code"),
      creditDays: find("Credit Days", "Credit Period"),
      invoiceDate: find("Dated", "Invoice Date", "Date"),
      invoiceNumber: find("Vrno", "Invoice No", "Invoice Number", "Voucher No"),
      transactionType: find("Vrtype", "Voucher Type", "Transaction Type"),
      balance: find("Bal", "Balance", "Outstanding", "Outstanding Amount"),
      ageDays: find("Day", "Days", "Age", "Age Days"),
    }
  }
  return {
    partyCode: find("Party Code", "A/C Code", "Account Code", "Customer Code"),
    name: find("Name", "Customer Name", "Account Name", "Party Name"),
    email: find("Email", "Email Address", "E-mail"),
  }
}

export function requiredMappingComplete(mapping: DuesMapping | MasterMapping): boolean {
  if ("accountCode" in mapping) {
    return Boolean(mapping.accountCode && mapping.creditDays && mapping.invoiceDate && mapping.invoiceNumber && mapping.transactionType && mapping.balance)
  }
  return Boolean(mapping.partyCode && mapping.name && mapping.email)
}

export function formatCurrency(value: number): string {
  return new Intl.NumberFormat("en-IN", {
    style: "currency",
    currency: "INR",
    maximumFractionDigits: 2,
  }).format(value)
}

export function formatDate(date: Date): string {
  return new Intl.DateTimeFormat("en-IN", { day: "2-digit", month: "short", year: "numeric", timeZone: "UTC" }).format(date)
}

export function formatIsoDate(date: Date): string {
  return date.toISOString().slice(0, 10)
}

export function renderTemplate(template: string, customer: CustomerReminder, stage: string): string {
  const invoiceRows = customer.invoices
    .map((invoice) => `${invoice.invoiceNumber} — ${formatCurrency(invoice.balance)} — due ${formatDate(invoice.dueDate)}`)
    .join("\n")
  return template
    .replaceAll("{{customer_name}}", customer.customerName)
    .replaceAll("{{party_code}}", customer.partyCode)
    .replaceAll("{{invoice_count}}", String(customer.invoices.length))
    .replaceAll("{{total_due}}", formatCurrency(customer.totalDue))
    .replaceAll("{{invoice_list}}", invoiceRows)
    .replaceAll("{{reminder_stage}}", stage)
}

export function plainTextToHtml(text: string): string {
  return text
    .split("\n")
    .map((line) => `<p style="margin:0 0 12px">${escapeHtml(line)}</p>`)
    .join("")
}

function escapeHtml(value: string): string {
  return value.replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;" })[char]!)
}

function startOfDay(date: Date): Date {
  return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()))
}

function addDays(date: Date, days: number): Date {
  return new Date(startOfDay(date).getTime() + days * DAY_MS)
}

function differenceInDays(later: Date, earlier: Date): number {
  return Math.round((startOfDay(later).getTime() - startOfDay(earlier).getTime()) / DAY_MS)
}

function normaliseHeader(value: unknown): string {
  return String(value ?? "").toLowerCase().replace(/[^a-z0-9]/g, "")
}

function normaliseCode(value: unknown): string {
  return String(value ?? "").trim().toUpperCase()
}

function cleanEmail(value: unknown): string | null {
  const email = String(value ?? "").trim().toLowerCase()
  return email && email !== "0" ? email : null
}

export function isValidEmail(email: string): boolean {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)
}

function stripLocation(value: string): string {
  return value.split(">>")[0].trim()
}
