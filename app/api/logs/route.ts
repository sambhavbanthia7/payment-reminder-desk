import { listLogs } from "@/lib/server/database"

export async function GET() {
  try {
    const rows = await listLogs()
    return Response.json({ logs: rows.map((row) => ({
      id: row.id,
      partyCode: row.party_code,
      customerName: row.customer_name,
      email: row.email,
      stage: row.reminder_stage,
      invoiceNumbers: JSON.parse(row.invoice_numbers),
      invoiceCount: row.invoice_count,
      totalDue: row.total_due,
      status: row.status,
      errorMessage: row.error_message ?? undefined,
      sentAt: row.sent_at,
    })) })
  } catch {
    return Response.json({ logs: [] })
  }
}
