"use client"

import { useEffect, useMemo, useRef, useState } from "react"
import {
  AlertTriangle, ArrowRight, BadgeIndianRupee, CalendarClock, Check, ChevronRight,
  CircleAlert, Clock3, FileSpreadsheet, LayoutDashboard, ListChecks, LogOut, Mail,
  MailCheck, RefreshCw, Send, Settings2, ShieldCheck, Upload,
} from "lucide-react"
import { toast } from "sonner"

import { Badge } from "@/components/ui/badge"
import { Button } from "@/components/ui/button"
import { Checkbox } from "@/components/ui/checkbox"
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle } from "@/components/ui/dialog"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select"
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table"
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs"
import { Textarea } from "@/components/ui/textarea"
import { Toaster } from "@/components/ui/sonner"
import {
  applyEmailOverrides, buildInvoices, deriveReportDate, detectMapping, formatCurrency, formatDate,
  formatIsoDate, groupCustomers, isValidEmail, plainTextToHtml, renderTemplate,
  type CampaignConfig, type CustomerReminder, type DuesMapping, type MasterMapping,
  type ReminderInvoice, type SendLog, type SheetTable,
} from "@/lib/reminders"
import { sampleDues, sampleMaster } from "@/lib/sample-data"
import { readWorkbook, selectBestSheet } from "@/lib/xlsx"

type Page = "overview" | "send" | "exceptions" | "logs"
type Bucket = "week" | "three" | "overdue" | "all"
type RowPreview = { title: string; invoices: ReminderInvoice[] } | null
type AuthStatus = { loading: boolean; configured: boolean; connected: boolean; name?: string; email?: string }

const defaultConfig: CampaignConfig = {
  firstThreshold: 7,
  secondThreshold: 3,
  defaultCreditDays: 30,
  asOfDate: "2026-08-21",
  subject: "Payment reminder: {{invoice_count}} invoice(s) due",
  body: `Dear {{customer_name}},

This is a reminder that the following invoice(s) remain outstanding:

{{invoice_list}}

Total outstanding: {{total_due}}

If payment has already been made, please share the transaction details so we can update our records.

Regards,
Accounts Receivable`,
}

const issueCopy = {
  unmatched: "Not found in master",
  missing_email: "Email missing",
  invalid_email: "Invalid email",
  negative_credit: "Invalid credit days",
}

export default function PaymentReminderDashboard() {
  const [auth, setAuth] = useState<AuthStatus>({ loading: true, configured: false, connected: false })
  const [demoMode, setDemoMode] = useState(false)
  const [page, setPage] = useState<Page>("overview")
  const [setupOpen, setSetupOpen] = useState(false)
  const [setupTab, setSetupTab] = useState("parameters")
  const [dues, setDues] = useState<SheetTable>(sampleDues)
  const [master, setMaster] = useState<SheetTable>(sampleMaster)
  const [duesMap, setDuesMap] = useState<DuesMapping>(() => detectMapping(sampleDues.headers, "dues"))
  const [masterMap, setMasterMap] = useState<MasterMapping>(() => detectMapping(sampleMaster.headers, "master"))
  const [config, setConfig] = useState(defaultConfig)
  const [bucket, setBucket] = useState<Bucket>("week")
  const [selected, setSelected] = useState<Set<string>>(new Set())
  const [logs, setLogs] = useState<SendLog[]>([])
  const [emailOverrides, setEmailOverrides] = useState<Record<string, string>>({})
  const [rowPreview, setRowPreview] = useState<RowPreview>(null)
  const [sending, setSending] = useState(false)
  const duesInput = useRef<HTMLInputElement>(null)
  const masterInput = useRef<HTMLInputElement>(null)

  useEffect(() => {
    // This session flag is an external browser value and must be read after hydration.
    // eslint-disable-next-line react-hooks/set-state-in-effect
    setDemoMode(sessionStorage.getItem("payment-reminder-demo") === "1")
    fetch("/api/auth/microsoft/status")
      .then((response) => response.json())
      .then((data) => setAuth({ loading: false, ...data }))
      .catch(() => setAuth({ loading: false, configured: false, connected: false }))
  }, [])

  useEffect(() => {
    if (!auth.connected && !demoMode) return
    fetch("/api/logs")
      .then((response) => response.ok ? response.json() : Promise.reject())
      .then((data) => setLogs(data.logs ?? []))
      .catch(() => {
        const stored = localStorage.getItem("payment-reminder-logs")
        if (stored) setLogs(JSON.parse(stored))
      })
  }, [auth.connected, demoMode])

  const sourceInvoices = useMemo(
    () => buildInvoices(dues, master, duesMap, masterMap, config),
    [dues, master, duesMap, masterMap, config],
  )
  const invoices = useMemo(
    () => applyEmailOverrides(sourceInvoices, emailOverrides),
    [sourceInvoices, emailOverrides],
  )
  const validInvoices = invoices.filter((invoice) => invoice.issue !== "negative_credit")
  const totalDue = validInvoices.reduce((sum, invoice) => sum + invoice.balance, 0)
  const weekInvoices = validInvoices.filter((invoice) => invoice.daysRemaining >= 0 && invoice.daysRemaining <= config.firstThreshold)
  const threeDayInvoices = validInvoices.filter((invoice) => invoice.daysRemaining >= 0 && invoice.daysRemaining <= config.secondThreshold)
  const overdueInvoices = validInvoices.filter((invoice) => invoice.daysRemaining < 0)
  const exceptions = groupCustomers(invoices.filter((invoice) => invoice.issue || invoice.usedDefaultCredit))
  const sentLogs = logs.filter((log) => log.status === "sent")

  const bucketInvoices = bucket === "week" ? weekInvoices : bucket === "three" ? threeDayInvoices : bucket === "overdue" ? overdueInvoices : validInvoices
  const bucketCustomers = groupCustomers(bucketInvoices)
  const sendableCustomers = bucketCustomers.filter((customer) => !customer.issue)

  const enterDemo = () => {
    sessionStorage.setItem("payment-reminder-demo", "1")
    setDemoMode(true)
  }

  if (auth.loading) return <LoadingScreen />
  if (!auth.connected && !demoMode) return <AuthGate configured={auth.configured} onPreview={enterDemo} />

  const toggleCustomer = (partyCode: string) => {
    setSelected((current) => {
      const next = new Set(current)
      if (next.has(partyCode)) next.delete(partyCode)
      else next.add(partyCode)
      return next
    })
  }

  const changeBucket = (value: Bucket) => {
    setBucket(value)
    setSelected(new Set())
  }

  const sendSelected = async () => {
    const customers = sendableCustomers.filter((customer) => selected.has(customer.partyCode))
    if (!customers.length) return
    setSending(true)
    const stage = bucket === "three" ? `${config.secondThreshold}-day priority` : bucket === "overdue" ? "Overdue" : `${config.firstThreshold}-day window`
    const messages = customers.map((customer) => ({
      campaignId: "test-receivables",
      partyCode: customer.partyCode,
      customerName: customer.customerName,
      email: customer.email!,
      subject: renderTemplate(config.subject, customer, stage),
      html: plainTextToHtml(renderTemplate(config.body, customer, stage)),
      stage,
      invoiceNumbers: customer.invoices.map((invoice) => invoice.invoiceNumber),
      invoiceCount: customer.invoices.length,
      totalDue: customer.totalDue,
    }))
    try {
      const response = await fetch("/api/outlook/send", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ messages, demo: demoMode }),
      })
      if (!response.ok) throw new Error((await response.json()).error ?? "Send failed")
      const data = await response.json()
      setLogs((current) => [...data.logs, ...current])
      setSelected(new Set())
      toast.success(demoMode ? "Email previews recorded. Connect Outlook to send." : `${data.logs.filter((log: SendLog) => log.status === "sent").length} email(s) sent`)
    } catch (error) {
      if (demoMode) {
        const previewLogs = messages.map<SendLog>((message) => ({ id: crypto.randomUUID(), ...message, status: "previewed", sentAt: new Date().toISOString() }))
        const next = [...previewLogs, ...logs]
        setLogs(next)
        localStorage.setItem("payment-reminder-logs", JSON.stringify(next))
        setSelected(new Set())
        toast.info("Email previews recorded. Connect Outlook to send.")
      } else toast.error(error instanceof Error ? error.message : "Emails could not be sent")
    } finally {
      setSending(false)
    }
  }

  const uploadWorkbook = async (file: File, kind: "dues" | "master") => {
    try {
      const workbook = await readWorkbook(file)
      const table = selectBestSheet(workbook, kind)
      if (kind === "dues") {
        const mapping = detectMapping(table.headers, "dues")
        setDues(table)
        setDuesMap(mapping)
        setConfig((current) => ({ ...current, asOfDate: formatIsoDate(deriveReportDate(table, mapping)) }))
      } else {
        setMaster(table)
        setMasterMap(detectMapping(table.headers, "master"))
        setEmailOverrides({})
      }
      setSetupTab("mapping")
      toast.success(`${file.name} loaded`)
    } catch (error) {
      toast.error(error instanceof Error ? error.message : "Workbook could not be read")
    }
  }

  const resetSample = () => {
    setDues(sampleDues); setMaster(sampleMaster)
    setDuesMap(detectMapping(sampleDues.headers, "dues")); setMasterMap(detectMapping(sampleMaster.headers, "master"))
    setConfig(defaultConfig); setEmailOverrides({}); toast.success("Test files restored")
  }

  return (
    <div className="min-h-screen bg-[#f4f6f8] text-[#17212b]">
      <Toaster position="top-right" />
      <aside className="fixed inset-y-0 left-0 z-30 hidden w-[248px] flex-col bg-[#0a192f] text-white lg:flex">
        <div className="flex h-20 items-center gap-3 border-b border-white/10 px-6">
          <div className="grid size-9 place-items-center rounded-lg bg-[#D9CEB4] text-[#0a192f]"><BadgeIndianRupee className="size-5" /></div>
          <div><p className="font-semibold tracking-tight">Payment Desk</p><p className="text-xs text-white/55">Accounts receivable</p></div>
        </div>
        <nav className="flex-1 space-y-1 px-3 py-6">
          <NavButton active={page === "overview"} icon={LayoutDashboard} label="Overview" onClick={() => setPage("overview")} />
          <NavButton active={page === "send"} icon={Send} label="Send reminders" onClick={() => setPage("send")} />
          <NavButton active={page === "exceptions"} icon={CircleAlert} label="Needs attention" count={exceptions.length} onClick={() => setPage("exceptions")} />
          <NavButton active={page === "logs"} icon={ListChecks} label="Email log" onClick={() => setPage("logs")} />
        </nav>
        <div className="border-t border-white/10 p-4">
          <div className="mb-3 flex items-center gap-3 rounded-lg bg-white/5 p-3">
            <div className="grid size-8 place-items-center rounded-full bg-white/10 text-xs font-semibold">{(auth.name ?? "FO").slice(0, 2).toUpperCase()}</div>
            <div className="min-w-0"><p className="truncate text-sm font-medium">{auth.name ?? "Finance operator"}</p><p className="truncate text-xs text-white/50">{demoMode ? "Sample workspace" : auth.email}</p></div>
          </div>
          {demoMode ? <button className="flex items-center gap-2 text-xs text-white/55 hover:text-white" onClick={() => { sessionStorage.removeItem("payment-reminder-demo"); location.reload() }}><LogOut className="size-3.5" /> Exit sample workspace</button> : <a href="/api/auth/microsoft/logout" className="flex items-center gap-2 text-xs text-white/55 hover:text-white"><LogOut className="size-3.5" /> Sign out</a>}
        </div>
      </aside>

      <main className="min-h-screen lg:pl-[248px]">
        <header className="sticky top-0 z-20 flex h-20 items-center justify-between border-b border-[#dfe4e7] bg-white/95 px-5 backdrop-blur md:px-8">
          <div><div className="mb-1 flex items-center gap-2 text-xs font-medium text-[#6b7780]"><span>Campaigns</span><ChevronRight className="size-3" /><span>August receivables</span></div><h1 className="text-lg font-semibold tracking-tight">{pageTitle(page)}</h1></div>
          <div className="flex items-center gap-2"><Badge variant="outline" className={demoMode ? "border-amber-200 bg-amber-50 text-amber-800" : "border-emerald-200 bg-emerald-50 text-emerald-800"}><span className={`size-1.5 rounded-full ${demoMode ? "bg-amber-500" : "bg-emerald-500"}`} />{demoMode ? "Sample mode" : "Outlook connected"}</Badge><Button variant="outline" onClick={() => setSetupOpen(true)}><Settings2 /> Campaign setup</Button></div>
        </header>

        <div className="mx-auto max-w-[1440px] p-5 md:p-8">
          {page === "overview" && <Overview config={config} totalDue={totalDue} allInvoices={validInvoices} weekInvoices={weekInvoices} threeDayInvoices={threeDayInvoices} overdueInvoices={overdueInvoices} logs={sentLogs} exceptions={exceptions} customers={groupCustomers(weekInvoices)} onPreview={(title, previewInvoices) => setRowPreview({ title, invoices: previewInvoices })} onSend={() => setPage("send")} onExceptions={() => setPage("exceptions")} />}
          {page === "send" && <SendCenter bucket={bucket} onBucket={changeBucket} customers={bucketCustomers} selected={selected} onToggle={toggleCustomer} onSelectAll={() => setSelected(new Set(sendableCustomers.map((customer) => customer.partyCode)))} onClear={() => setSelected(new Set())} onSend={sendSelected} sending={sending} demoMode={demoMode} config={config} />}
          {page === "exceptions" && <Exceptions customers={exceptions} defaultCreditDays={config.defaultCreditDays} overrides={emailOverrides} onEmailOverride={(partyCode, email) => setEmailOverrides((current) => ({ ...current, [partyCode]: email }))} />}
          {page === "logs" && <Logs logs={logs} />}
        </div>
      </main>

      <CampaignSetup open={setupOpen} onOpenChange={setSetupOpen} tab={setupTab} setTab={setSetupTab} config={config} setConfig={setConfig} dues={dues} master={master} duesMap={duesMap} setDuesMap={setDuesMap} masterMap={masterMap} setMasterMap={setMasterMap} duesInput={duesInput} masterInput={masterInput} onUpload={uploadWorkbook} onReset={resetSample} />
      <RowsPreviewDialog preview={rowPreview} onOpenChange={(open) => !open && setRowPreview(null)} duesMap={duesMap} defaultCreditDays={config.defaultCreditDays} />
    </div>
  )
}

function AuthGate({ configured, onPreview }: { configured: boolean; onPreview: () => void }) {
  const authError = typeof window !== "undefined" ? new URLSearchParams(window.location.search).get("auth_error") : null
  return (
    <main className="flex min-h-screen flex-col items-center justify-center bg-[#f4f6f8] p-6">
      <div className="mb-10 text-center">
        <img src="https://jalbath.com/wp-content/uploads/2025/12/header-logo.svg" alt="Jalbat Logo" className="mx-auto h-16 object-contain drop-shadow-md" />
        <h1 className="mt-6 text-3xl font-semibold tracking-tight text-[#17212b]">Payment Reminder Desk</h1>
      </div>

      <div className="w-full max-w-md rounded-2xl border border-[#dfe7e4] bg-white p-8 shadow-[0_20px_60px_rgba(10,25,47,.08)]">
        {authError && <div className="mb-4 rounded-lg border border-red-200 bg-red-50 p-3 text-sm text-red-800">Outlook sign-in could not be completed. Please try again.</div>}
        {configured ? (
          <Button asChild size="lg" className="h-12 w-full bg-[#0a192f] hover:bg-[#061020] text-[#D9CEB4]">
            <a href="/api/auth/microsoft/start"><Mail className="mr-2 size-5" /> Sign in</a>
          </Button>
        ) : (
          <div className="rounded-lg border border-amber-200 bg-amber-50 p-3 text-sm leading-5 text-amber-900">
            Outlook credentials have not been added to this deployment yet. The sample workspace is fully available for review.
          </div>
        )}
        
        <div className="my-6 flex items-center gap-3 text-xs text-[#8a949a]"><span className="h-px flex-1 bg-[#e5e9eb]" />or<span className="h-px flex-1 bg-[#e5e9eb]" /></div>
        
        <Button variant="outline" size="lg" className="h-12 w-full border-[#0a192f] text-[#0a192f] hover:bg-[#f4f6f8]" onClick={onPreview}>
          <FileSpreadsheet className="mr-2 size-5" /> Review/Test Page
        </Button>
      </div>
    </main>
  )
}

function Overview({ config, totalDue, allInvoices, weekInvoices, threeDayInvoices, overdueInvoices, logs, exceptions, customers, onPreview, onSend, onExceptions }: { config: CampaignConfig; totalDue: number; allInvoices: ReminderInvoice[]; weekInvoices: ReminderInvoice[]; threeDayInvoices: ReminderInvoice[]; overdueInvoices: ReminderInvoice[]; logs: SendLog[]; exceptions: CustomerReminder[]; customers: CustomerReminder[]; onPreview: (title: string, invoices: ReminderInvoice[]) => void; onSend: () => void; onExceptions: () => void }) {
  const weekTotal = weekInvoices.reduce((sum, row) => sum + row.balance, 0)
  const threeTotal = threeDayInvoices.reduce((sum, row) => sum + row.balance, 0)
  const overdueTotal = overdueInvoices.reduce((sum, row) => sum + row.balance, 0)
  const allCustomerCount = groupCustomers(allInvoices).length
  const weekCustomerCount = groupCustomers(weekInvoices).length
  const threeCustomerCount = groupCustomers(threeDayInvoices).length
  const overdueCustomerCount = groupCustomers(overdueInvoices).length
  return <div className="space-y-6">
    <section className="flex flex-col justify-between gap-4 rounded-xl border border-[#dfe5e7] bg-white p-5 md:flex-row md:items-center"><div className="flex items-start gap-4"><div className="grid size-10 place-items-center rounded-lg bg-[#e9f5ef] text-[#1e6b54]"><Check /></div><div><p className="font-semibold">Test campaign is ready</p><p className="mt-1 text-sm text-[#68757d]">Report date {formatDate(new Date(`${config.asOfDate}T00:00:00Z`))} · {config.defaultCreditDays}-day placeholder applied where credit days are blank</p></div></div><Button onClick={onSend} className="bg-[#0a192f] hover:bg-[#061020] text-[#D9CEB4]">Review recipients <ArrowRight className="ml-2" /></Button></section>
    <section className="grid gap-4 xl:grid-cols-4"><MetricCard icon={BadgeIndianRupee} label="Total open invoices" value={formatCurrency(totalDue)} detail={`${allInvoices.length} invoices · ${allCustomerCount} customers`} tone="slate" onClick={() => onPreview("All open invoice rows", allInvoices)} /><MetricCard icon={CalendarClock} label={`Due within ${config.firstThreshold} days`} value={formatCurrency(weekTotal)} detail={`${weekInvoices.length} invoices · ${weekCustomerCount} customers · ${logs.filter((log) => log.stage.includes(`${config.firstThreshold}-day`)).length} emails sent`} tone="green" onClick={() => onPreview(`Rows due within ${config.firstThreshold} days`, weekInvoices)} /><MetricCard icon={Clock3} label={`Due within ${config.secondThreshold} days`} value={formatCurrency(threeTotal)} detail={`${threeDayInvoices.length} invoices · ${threeCustomerCount} customers · ${logs.filter((log) => log.stage.includes(`${config.secondThreshold}-day`)).length} emails sent`} tone="amber" onClick={() => onPreview(`Rows due within ${config.secondThreshold} days`, threeDayInvoices)} /><MetricCard icon={AlertTriangle} label="Past due" value={formatCurrency(overdueTotal)} detail={`${overdueInvoices.length} invoices · ${overdueCustomerCount} customers`} tone="red" onClick={() => onPreview("Past-due invoice rows", overdueInvoices)} /></section>
    <section className="grid gap-6 xl:grid-cols-[1.55fr_.85fr]">
      <div className="rounded-xl border border-[#dfe5e7] bg-white"><div className="flex items-center justify-between border-b border-[#e7ebed] px-5 py-4"><div><h2 className="font-semibold">Due this week</h2><p className="mt-1 text-xs text-[#748087]">Customer totals from invoices due in the configured window</p></div><Button variant="ghost" size="sm" onClick={onSend}>View all <ChevronRight /></Button></div><Table><TableHeader><TableRow><TableHead>Customer</TableHead><TableHead>Invoices</TableHead><TableHead>Nearest due</TableHead><TableHead className="text-right">Amount</TableHead><TableHead>Status</TableHead></TableRow></TableHeader><TableBody>{customers.slice(0, 7).map((customer) => <TableRow key={customer.partyCode}><TableCell><p className="font-medium">{customer.customerName}</p><p className="text-xs text-[#78848b]">{customer.partyCode}</p></TableCell><TableCell>{customer.invoices.length}</TableCell><TableCell>{dayLabel(customer.nearestDue)}</TableCell><TableCell className="text-right font-medium">{formatCurrency(customer.totalDue)}</TableCell><TableCell>{customer.issue ? <IssueBadge issue={customer.issue} /> : <Badge variant="outline" className="border-emerald-200 bg-emerald-50 text-emerald-800">Ready</Badge>}</TableCell></TableRow>)}</TableBody></Table></div>
      <div className="rounded-xl border border-[#dfe5e7] bg-white p-5"><div className="flex items-start justify-between"><div><h2 className="font-semibold">Needs attention</h2><p className="mt-1 text-xs text-[#748087]">Fix these before sending</p></div><span className="rounded-full bg-red-50 px-2.5 py-1 text-xs font-semibold text-red-700">{exceptions.length}</span></div><div className="mt-5 space-y-3">{exceptions.slice(0, 5).map((customer) => <button key={customer.partyCode} onClick={onExceptions} className="flex w-full items-center gap-3 rounded-lg border border-[#e6eaec] p-3 text-left hover:border-[#bcc9c8] hover:bg-[#fafcfc]"><div className="grid size-8 shrink-0 place-items-center rounded-md bg-amber-50 text-amber-700"><CircleAlert className="size-4" /></div><div className="min-w-0 flex-1"><p className="truncate text-sm font-medium">{customer.customerName}</p><p className="text-xs text-[#7b868d]">{customer.issue ? issueCopy[customer.issue] : `Using ${config.defaultCreditDays}-day placeholder`}</p></div><ChevronRight className="size-4 text-[#9aa3a8]" /></button>)}</div><Button variant="outline" className="mt-4 w-full" onClick={onExceptions}>Review all exceptions</Button></div>
    </section>
  </div>
}

function SendCenter({ bucket, onBucket, customers, selected, onToggle, onSelectAll, onClear, onSend, sending, demoMode, config }: { bucket: Bucket; onBucket: (bucket: Bucket) => void; customers: CustomerReminder[]; selected: Set<string>; onToggle: (code: string) => void; onSelectAll: () => void; onClear: () => void; onSend: () => void; sending: boolean; demoMode: boolean; config: CampaignConfig }) {
  const selectedCustomers = customers.filter((customer) => selected.has(customer.partyCode))
  const selectedTotal = selectedCustomers.reduce((sum, customer) => sum + customer.totalDue, 0)
  return <div className="space-y-5">
    <section className="flex flex-col justify-between gap-4 rounded-xl border border-[#dfe5e7] bg-white p-5 lg:flex-row lg:items-center"><div><h2 className="font-semibold">Choose who receives a reminder</h2><p className="mt-1 text-sm text-[#68757d]">Customers with missing or invalid contact data remain visible but cannot be selected.</p></div><Select value={bucket} onValueChange={(value) => onBucket(value as Bucket)}><SelectTrigger className="w-[230px]"><SelectValue /></SelectTrigger><SelectContent><SelectItem value="week">Due within {config.firstThreshold} days</SelectItem><SelectItem value="three">Due within {config.secondThreshold} days</SelectItem><SelectItem value="overdue">Past due</SelectItem><SelectItem value="all">All open invoices</SelectItem></SelectContent></Select></section>
    <section className="overflow-hidden rounded-xl border border-[#dfe5e7] bg-white"><div className="flex items-center justify-between border-b border-[#e7ebed] px-5 py-4"><p className="text-sm"><span className="font-semibold">{customers.length}</span> customers in this view</p><div className="flex gap-2"><Button variant="ghost" size="sm" onClick={onClear}>Clear</Button><Button variant="outline" size="sm" onClick={onSelectAll}>Select all ready</Button></div></div><Table><TableHeader><TableRow><TableHead className="w-10"><span className="sr-only">Select</span></TableHead><TableHead>Customer</TableHead><TableHead>Email</TableHead><TableHead>Invoices</TableHead><TableHead>Timing</TableHead><TableHead className="text-right">Outstanding</TableHead><TableHead>Readiness</TableHead></TableRow></TableHeader><TableBody>{customers.map((customer) => <TableRow key={customer.partyCode} data-state={selected.has(customer.partyCode) ? "selected" : undefined}><TableCell><Checkbox checked={selected.has(customer.partyCode)} disabled={Boolean(customer.issue)} onCheckedChange={() => onToggle(customer.partyCode)} aria-label={`Select ${customer.customerName}`} /></TableCell><TableCell><p className="font-medium">{customer.customerName}</p><p className="text-xs text-[#78848b]">{customer.partyCode}</p></TableCell><TableCell className="max-w-[220px] truncate">{customer.email ?? "—"}</TableCell><TableCell>{customer.invoices.length}</TableCell><TableCell>{dayLabel(customer.nearestDue)}</TableCell><TableCell className="text-right font-medium">{formatCurrency(customer.totalDue)}</TableCell><TableCell>{customer.issue ? <IssueBadge issue={customer.issue} /> : <Badge variant="outline" className="border-emerald-200 bg-emerald-50 text-emerald-800">Ready</Badge>}</TableCell></TableRow>)}</TableBody></Table>{!customers.length && <div className="grid place-items-center px-5 py-16 text-center"><MailCheck className="mb-3 size-8 text-[#93a09f]" /><p className="font-medium">No customers in this window</p><p className="mt-1 text-sm text-[#748087]">Choose another due-date bucket.</p></div>}</section>
    <section className="sticky bottom-4 flex flex-col justify-between gap-4 rounded-xl border border-[#b9c9c5] bg-[#0a192f] p-4 text-white shadow-xl md:flex-row md:items-center"><div className="flex items-center gap-5"><div><p className="text-xs text-white/55">Selected</p><p className="font-semibold">{selectedCustomers.length} customers</p></div><div className="h-8 w-px bg-white/15" /><div><p className="text-xs text-white/55">Invoice value</p><p className="font-semibold">{formatCurrency(selectedTotal)}</p></div></div><Button disabled={!selectedCustomers.length || sending} onClick={onSend} className="bg-[#D9CEB4] text-[#0a192f] hover:bg-[#C2B79F]">{sending ? <RefreshCw className="animate-spin" /> : <Send className="mr-2" />}{demoMode ? "Preview selected emails" : "Send selected emails"}</Button></section>
  </div>
}

function Exceptions({ customers, defaultCreditDays, overrides, onEmailOverride }: { customers: CustomerReminder[]; defaultCreditDays: number; overrides: Record<string, string>; onEmailOverride: (partyCode: string, email: string) => void }) {
  return <div className="space-y-5"><div className="rounded-xl border border-amber-200 bg-amber-50 p-4 text-sm text-amber-950"><p className="font-medium">These records need manual attention.</p><p className="mt-1 text-amber-800">Enter a valid replacement email to make missing, invalid, or unmatched customers ready for this campaign. The uploaded master file is not changed. Missing credit days use the current {defaultCreditDays}-day placeholder and do not block reminders.</p></div><div className="overflow-x-auto rounded-xl border border-[#dfe5e7] bg-white"><Table><TableHeader><TableRow><TableHead>Customer</TableHead><TableHead>Issue</TableHead><TableHead className="min-w-[310px]">Email / campaign override</TableHead><TableHead>Invoices affected</TableHead><TableHead className="text-right">Value</TableHead><TableHead>Suggested follow-up</TableHead></TableRow></TableHeader><TableBody>{customers.map((customer) => <TableRow key={customer.partyCode}><TableCell><p className="font-medium">{customer.customerName}</p><p className="text-xs text-[#78848b]">{customer.partyCode}</p></TableCell><TableCell>{customer.issue ? <IssueBadge issue={customer.issue} /> : <Badge variant="outline" className="border-blue-200 bg-blue-50 text-blue-800">Default credit used</Badge>}</TableCell><TableCell>{customer.issue === "unmatched" || customer.issue === "missing_email" || customer.issue === "invalid_email" ? <EmailOverrideEditor customer={customer} savedValue={overrides[customer.partyCode]} onSave={onEmailOverride} /> : <span className="text-[#5f6c73]">{customer.email ?? "—"}</span>}</TableCell><TableCell>{customer.invoices.length}</TableCell><TableCell className="text-right font-medium">{formatCurrency(customer.totalDue)}</TableCell><TableCell className="text-[#5f6c73]">{customer.issue === "unmatched" ? "Enter a known billing email" : customer.issue === "missing_email" ? "Enter the billing email" : customer.issue === "invalid_email" ? "Correct the email format" : customer.issue === "negative_credit" ? "Confirm the agreed credit term" : `Confirm ${defaultCreditDays}-day term`}</TableCell></TableRow>)}</TableBody></Table></div>{!customers.length && <div className="rounded-xl border border-emerald-200 bg-emerald-50 p-6 text-center text-emerald-900"><MailCheck className="mx-auto mb-2 size-6" /><p className="font-medium">No records need attention</p></div>}</div>
}

function EmailOverrideEditor({ customer, savedValue, onSave }: { customer: CustomerReminder; savedValue?: string; onSave: (partyCode: string, email: string) => void }) {
  const [draft, setDraft] = useState(savedValue ?? customer.email ?? "")
  const valid = isValidEmail(draft.trim())
  const save = () => {
    if (!valid) return toast.error("Enter a valid email address")
    onSave(customer.partyCode, draft.trim().toLowerCase())
    toast.success(`Email saved for ${customer.customerName}`)
  }
  return <div><div className="flex gap-2"><Input type="email" value={draft} onChange={(event) => setDraft(event.target.value)} onKeyDown={(event) => { if (event.key === "Enter") save() }} placeholder="billing@customer.com" aria-label={`Email override for ${customer.customerName}`} className={!draft || valid ? "" : "border-red-300 focus-visible:ring-red-200"} /><Button size="sm" variant="outline" disabled={!valid} onClick={save}>Save</Button></div><p className={`mt-1 text-[11px] ${draft && !valid ? "text-red-700" : "text-[#7b868d]"}`}>{draft && !valid ? "Enter a complete email address" : "Used for this campaign only"}</p></div>
}

function RowsPreviewDialog({ preview, onOpenChange, duesMap, defaultCreditDays }: { preview: RowPreview; onOpenChange: (open: boolean) => void; duesMap: DuesMapping; defaultCreditDays: number }) {
  const invoices = preview?.invoices ?? []
  const total = invoices.reduce((sum, invoice) => sum + invoice.balance, 0)
  const customerCount = groupCustomers(invoices).length
  return <Dialog open={Boolean(preview)} onOpenChange={onOpenChange}><DialogContent className="max-h-[92vh] gap-0 overflow-hidden p-0 sm:max-w-[min(96vw,1400px)]"><DialogHeader className="border-b border-[#e5e9eb] p-6 pr-14"><DialogTitle>{preview?.title ?? "Invoice rows"}</DialogTitle><DialogDescription>Exact uploaded dues rows included in this dashboard total.</DialogDescription><div className="flex flex-wrap gap-2 pt-2 text-xs"><Badge variant="outline">{invoices.length} invoices</Badge><Badge variant="outline">{customerCount} customers</Badge><Badge variant="outline">{formatCurrency(total)}</Badge></div></DialogHeader><div className="max-h-[68vh] overflow-auto"><Table className="min-w-[1280px]"><TableHeader className="sticky top-0 z-10 bg-white"><TableRow><TableHead>Excel row</TableHead><TableHead>Account name</TableHead><TableHead>A/C code</TableHead><TableHead>Invoice date</TableHead><TableHead>Invoice no.</TableHead><TableHead>Credit days</TableHead><TableHead className="text-right">Raw balance</TableHead><TableHead>Calculated due</TableHead><TableHead>Timing</TableHead><TableHead>Email status</TableHead></TableRow></TableHeader><TableBody>{invoices.map((invoice) => <TableRow key={`${invoice.id}-${invoice.sourceRowNumber}`}><TableCell className="font-mono text-xs">{invoice.sourceRowNumber}</TableCell><TableCell>{rawCellValue(invoice.sourceRow[duesMap.accountName]) || invoice.customerName}</TableCell><TableCell className="font-mono text-xs">{rawCellValue(invoice.sourceRow[duesMap.accountCode]) || invoice.partyCode}</TableCell><TableCell>{rawCellValue(invoice.sourceRow[duesMap.invoiceDate])}</TableCell><TableCell className="font-medium">{invoice.invoiceNumber}</TableCell><TableCell>{invoice.usedDefaultCredit ? <span>{defaultCreditDays} <span className="text-xs text-amber-700">(blank → default)</span></span> : invoice.creditDays}</TableCell><TableCell className="text-right">{rawCellValue(invoice.sourceRow[duesMap.balance])}</TableCell><TableCell>{formatDate(invoice.dueDate)}</TableCell><TableCell>{dayLabel(invoice.daysRemaining)}</TableCell><TableCell>{invoice.issue ? <IssueBadge issue={invoice.issue} /> : <Badge variant="outline" className="border-emerald-200 bg-emerald-50 text-emerald-800">Ready</Badge>}</TableCell></TableRow>)}</TableBody></Table>{!invoices.length && <div className="grid place-items-center px-6 py-16 text-center"><FileSpreadsheet className="mb-3 size-8 text-[#93a09f]" /><p className="font-medium">No Excel rows in this bucket</p></div>}</div><DialogFooter className="border-t border-[#e5e9eb] p-4"><Button variant="outline" onClick={() => onOpenChange(false)}>Close preview</Button></DialogFooter></DialogContent></Dialog>
}

function Logs({ logs }: { logs: SendLog[] }) {
  return <div className="overflow-hidden rounded-xl border border-[#dfe5e7] bg-white"><div className="border-b border-[#e7ebed] px-5 py-4"><h2 className="font-semibold">Reminder activity</h2><p className="mt-1 text-xs text-[#748087]">A record of every attempted send or sample preview</p></div>{logs.length ? <Table><TableHeader><TableRow><TableHead>Date & time</TableHead><TableHead>Customer</TableHead><TableHead>Email</TableHead><TableHead>Stage</TableHead><TableHead>Invoices</TableHead><TableHead className="text-right">Amount</TableHead><TableHead>Status</TableHead></TableRow></TableHeader><TableBody>{logs.map((log) => <TableRow key={log.id}><TableCell>{new Date(log.sentAt).toLocaleString("en-IN", { dateStyle: "medium", timeStyle: "short" })}</TableCell><TableCell><p className="font-medium">{log.customerName}</p><p className="text-xs text-[#78848b]">{log.partyCode}</p></TableCell><TableCell>{log.email}</TableCell><TableCell>{log.stage}</TableCell><TableCell>{log.invoiceCount}</TableCell><TableCell className="text-right font-medium">{formatCurrency(log.totalDue)}</TableCell><TableCell><Badge variant="outline" className={log.status === "sent" ? "border-emerald-200 bg-emerald-50 text-emerald-800" : log.status === "failed" ? "border-red-200 bg-red-50 text-red-800" : "border-blue-200 bg-blue-50 text-blue-800"}>{log.status}</Badge></TableCell></TableRow>)}</TableBody></Table> : <div className="grid place-items-center px-5 py-20 text-center"><ListChecks className="mb-3 size-9 text-[#a2acab]" /><p className="font-medium">No reminder activity yet</p><p className="mt-1 text-sm text-[#748087]">Sent emails and sample previews will appear here.</p></div>}</div>
}

function CampaignSetup({ open, onOpenChange, tab, setTab, config, setConfig, dues, master, duesMap, setDuesMap, masterMap, setMasterMap, duesInput, masterInput, onUpload, onReset }: { open: boolean; onOpenChange: (open: boolean) => void; tab: string; setTab: (tab: string) => void; config: CampaignConfig; setConfig: React.Dispatch<React.SetStateAction<CampaignConfig>>; dues: SheetTable; master: SheetTable; duesMap: DuesMapping; setDuesMap: React.Dispatch<React.SetStateAction<DuesMapping>>; masterMap: MasterMapping; setMasterMap: React.Dispatch<React.SetStateAction<MasterMapping>>; duesInput: React.RefObject<HTMLInputElement | null>; masterInput: React.RefObject<HTMLInputElement | null>; onUpload: (file: File, kind: "dues" | "master") => void; onReset: () => void }) {
  return <Dialog open={open} onOpenChange={onOpenChange}><DialogContent className="max-h-[90vh] overflow-y-auto sm:max-w-4xl"><DialogHeader><DialogTitle>Configure reminder campaign</DialogTitle><DialogDescription>Set the reminder rules, load the latest exports, and confirm how columns should be interpreted.</DialogDescription></DialogHeader><Tabs value={tab} onValueChange={setTab} className="mt-2"><TabsList className="grid w-full grid-cols-3"><TabsTrigger value="parameters">1. Parameters</TabsTrigger><TabsTrigger value="files">2. Files</TabsTrigger><TabsTrigger value="mapping">3. Column mapping</TabsTrigger></TabsList>
    <TabsContent value="parameters" className="space-y-5 pt-4"><div className="grid gap-4 md:grid-cols-3"><NumberField label="First reminder window" suffix="days" value={config.firstThreshold} onChange={(value) => setConfig((current) => ({ ...current, firstThreshold: value }))} /><NumberField label="Priority reminder window" suffix="days" value={config.secondThreshold} onChange={(value) => setConfig((current) => ({ ...current, secondThreshold: value }))} /><NumberField label="Blank credit-days placeholder" suffix="days" value={config.defaultCreditDays} onChange={(value) => setConfig((current) => ({ ...current, defaultCreditDays: value }))} /></div><div><Label htmlFor="report-date">Report date</Label><Input id="report-date" type="date" className="mt-2 max-w-xs" value={config.asOfDate} onChange={(event) => setConfig((current) => ({ ...current, asOfDate: event.target.value }))} /><p className="mt-1.5 text-xs text-[#748087]">Detected from Invoice Date + Day in the dues report. Change it if the export was generated on another date.</p></div><div><Label htmlFor="subject">Email subject</Label><Input id="subject" className="mt-2" value={config.subject} onChange={(event) => setConfig((current) => ({ ...current, subject: event.target.value }))} /></div><div><Label htmlFor="body">Email body</Label><Textarea id="body" className="mt-2 min-h-56 font-mono text-xs leading-5" value={config.body} onChange={(event) => setConfig((current) => ({ ...current, body: event.target.value }))} /><p className="mt-2 text-xs text-[#748087]">Available fields: {"{{customer_name}}"}, {"{{party_code}}"}, {"{{invoice_count}}"}, {"{{total_due}}"}, {"{{invoice_list}}"}</p></div></TabsContent>
    <TabsContent value="files" className="space-y-4 pt-4"><UploadCard title="Dues analysis" file={dues} onChoose={() => duesInput.current?.click()} /><UploadCard title="Customer master" file={master} onChoose={() => masterInput.current?.click()} /><input ref={duesInput} hidden type="file" accept=".xlsx" onChange={(event) => event.target.files?.[0] && onUpload(event.target.files[0], "dues")} /><input ref={masterInput} hidden type="file" accept=".xlsx" onChange={(event) => event.target.files?.[0] && onUpload(event.target.files[0], "master")} /><Button variant="ghost" size="sm" onClick={onReset}><RefreshCw /> Restore attached test files</Button></TabsContent>
    <TabsContent value="mapping" className="space-y-6 pt-4"><MappingSection title={`Dues · ${dues.sheetName}`} headers={dues.headers} fields={[["accountName","Account name"],["accountCode","Account / party code"],["creditDays","Credit days"],["invoiceDate","Invoice date"],["invoiceNumber","Invoice number"],["transactionType","Transaction type"],["balance","Outstanding balance"],["ageDays","Age / day (optional)"]]} mapping={duesMap} onChange={(key, value) => setDuesMap((current) => ({ ...current, [key]: value }))} /><MappingSection title={`Master · ${master.sheetName}`} headers={master.headers} fields={[["partyCode","Party code"],["name","Customer name"],["email","Email address"]]} mapping={masterMap} onChange={(key, value) => setMasterMap((current) => ({ ...current, [key]: value }))} /></TabsContent>
  </Tabs><DialogFooter><Button variant="outline" onClick={() => onOpenChange(false)}>Cancel</Button><Button onClick={() => { onOpenChange(false); toast.success("Campaign settings saved") }} className="bg-[#0a192f] hover:bg-[#061020] text-[#D9CEB4]">Save campaign</Button></DialogFooter></DialogContent></Dialog>
}

function MappingSection({ title, headers, fields, mapping, onChange }: { title: string; headers: string[]; fields: string[][]; mapping: DuesMapping | MasterMapping; onChange: (key: string, value: string) => void }) {
  const values = mapping as unknown as Record<string, string>
  return <div><div className="mb-3 flex items-center justify-between"><h3 className="font-medium">{title}</h3><Badge variant="outline">{headers.length} columns found</Badge></div><div className="grid gap-3 rounded-lg border border-[#e4e9ea] p-4 md:grid-cols-2">{fields.map(([key, label]) => <div key={key}><Label>{label}</Label><Select value={values[key] || "__none"} onValueChange={(value) => onChange(key, value === "__none" ? "" : value)}><SelectTrigger className="mt-1.5 w-full"><SelectValue placeholder="Choose column" /></SelectTrigger><SelectContent><SelectItem value="__none">Not mapped</SelectItem>{headers.map((header) => <SelectItem value={header} key={header}>{header}</SelectItem>)}</SelectContent></Select></div>)}</div></div>
}

function UploadCard({ title, file, onChoose }: { title: string; file: SheetTable; onChoose: () => void }) { return <div className="flex flex-col justify-between gap-4 rounded-xl border border-[#dfe5e7] p-4 sm:flex-row sm:items-center"><div className="flex items-center gap-3"><div className="grid size-10 place-items-center rounded-lg bg-emerald-50 text-emerald-700"><FileSpreadsheet /></div><div><p className="font-medium">{title}</p><p className="text-sm text-[#68757d]">{file.fileName} · {file.sheetName} · {file.rows.length} rows</p></div></div><Button variant="outline" onClick={onChoose}><Upload /> Replace file</Button></div> }
function NumberField({ label, suffix, value, onChange }: { label: string; suffix: string; value: number; onChange: (value: number) => void }) { return <div><Label>{label}</Label><div className="relative mt-2"><Input type="number" min={0} value={value} onChange={(event) => onChange(Math.max(0, Number(event.target.value)))} className="pr-12" /><span className="absolute right-3 top-2 text-xs text-[#748087]">{suffix}</span></div></div> }
function MetricCard({ icon: Icon, label, value, detail, tone, onClick }: { icon: typeof Mail; label: string; value: string; detail: string; tone: "slate" | "green" | "amber" | "red"; onClick: () => void }) { const tones = { slate: "bg-slate-100 text-slate-700", green: "bg-emerald-50 text-emerald-700", amber: "bg-amber-50 text-amber-700", red: "bg-red-50 text-red-700" }; return <button type="button" onClick={onClick} className="group rounded-xl border border-[#dfe5e7] bg-white p-5 text-left transition hover:-translate-y-0.5 hover:border-[#aab9b7] hover:shadow-[0_10px_30px_rgba(16,43,43,.08)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#1e6b54] focus-visible:ring-offset-2"><div className="flex items-start justify-between"><div className={`mb-5 grid size-9 place-items-center rounded-lg ${tones[tone]}`}><Icon className="size-4.5" /></div><span className="flex items-center gap-1 text-[11px] font-medium text-[#82908f] transition group-hover:text-[#1e6b54]">View rows <ChevronRight className="size-3.5" /></span></div><p className="text-xs font-medium text-[#6f7b82]">{label}</p><p className="mt-2 text-2xl font-semibold tracking-[-0.025em]">{value}</p><p className="mt-2 text-xs leading-5 text-[#7c878d]">{detail}</p></button> }
function NavButton({ active, icon: Icon, label, count, onClick }: { active: boolean; icon: typeof Mail; label: string; count?: number; onClick: () => void }) { return <button onClick={onClick} className={`flex w-full items-center gap-3 rounded-lg px-3 py-2.5 text-sm transition ${active ? "bg-white/12 text-white" : "text-white/60 hover:bg-white/6 hover:text-white"}`}><Icon className="size-4" /><span className="flex-1 text-left">{label}</span>{count ? <span className="rounded-full bg-amber-300 px-1.5 py-0.5 text-[10px] font-bold text-[#3b3010]">{count}</span> : null}</button> }
function IssueBadge({ issue }: { issue: NonNullable<CustomerReminder["issue"]> }) { return <Badge variant="outline" className="border-red-200 bg-red-50 text-red-800">{issueCopy[issue]}</Badge> }
function LoadingScreen() { return <main className="grid min-h-screen place-items-center bg-[#eff3f2]"><div className="text-center"><RefreshCw className="mx-auto size-6 animate-spin text-[#1e6b54]" /><p className="mt-3 text-sm text-[#68757d]">Preparing payment workspace…</p></div></main> }
function dayLabel(days: number): string { return days < 0 ? `${Math.abs(days)}d overdue` : days === 0 ? "Due today" : `${days}d remaining` }
function pageTitle(page: Page): string { return page === "overview" ? "Campaign overview" : page === "send" ? "Send reminders" : page === "exceptions" ? "Needs attention" : "Email log" }
function rawCellValue(value: unknown): string { return value instanceof Date ? formatDate(value) : String(value ?? "").trim() }
