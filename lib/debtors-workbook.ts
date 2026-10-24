import { strFromU8, strToU8, unzipSync, zipSync } from "fflate"
import { calculateInvoices, defaultSettings, invoiceId, today, validDate, type WorkbookState, type Receipt } from "./debtors.ts"

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
const CONTENT = "http://schemas.openxmlformats.org/package/2006/content-types"
type Cell = { ref: string; value: string | number; formula: string; error: boolean }
type Sheet = { name: string; path: string; document: XMLDocument; rows: Map<number, Map<string, Cell>>; editRows?: Map<number,Element>; editCells?: Map<string,Element>; sheetData?: Element }
type Package = { files: Record<string, Uint8Array>; workbook: XMLDocument; rels: XMLDocument; content: XMLDocument; sheets: Sheet[]; epoch: number }
type Scalar = string | number | boolean | null
const definitions = {
  activity: { name: "Activity Log", headers: ["ID", "Timestamp", "Action", "Invoice key", "Party", "Invoice", "Detail", "Previous value", "New value"] },
  reminders: { name: "Reminder Log", headers: ["ID", "Timestamp", "Review date", "Type", "Party", "Recipient", "Invoice keys (JSON)", "Invoice numbers (JSON)", "Outstanding", "Stage", "Subject", "Message", "Status", "Sent at", "Sent by", "Sent to", "Send result"], required: 13 },
  followups: { name: "Followup State", headers: ["Invoice key", "Flagged", "Manual escalation", "On hold", "Note"] },
  confirmations: { name: "Balance Confirmations", headers: ["ID", "Timestamp", "Party", "Confirmation date", "Balance", "Status", "Reference", "Note"] },
  settings: { name: "Desk Settings", headers: ["Setting", "Value"] },
}
function parseXml(source: string): XMLDocument {
  const doc = new DOMParser().parseFromString(source, "application/xml")
  if (doc.getElementsByTagName("parsererror").length) throw new Error("The workbook contains invalid XML")
  return doc
}
const elements = (node: Document | Element, name: string) => [...node.getElementsByTagNameNS("*", name)]
const serialise = (doc: XMLDocument) => strToU8(new XMLSerializer().serializeToString(doc))
const textAt = (node: Element, name: string) => elements(node, name)[0]?.textContent ?? ""
function pathFor(target: string): string {
  const parts: string[] = []
  for (const p of (target.startsWith("/") ? target.slice(1) : `xl/${target}`).split("/")) { if (p === "..") parts.pop(); else if (p && p !== ".") parts.push(p) }
  return parts.join("/")
}
function openPackage(bytes: Uint8Array): Package {
  let total = 0
  const files = unzipSync(bytes, { filter: (entry) => { total += entry.originalSize; if (total > 100 * 1024 * 1024) throw new Error("The uncompressed workbook exceeds 100 MB"); return true } })
  const read = (path: string) => { if (!files[path]) throw new Error(`Workbook part is missing: ${path}`); return parseXml(strFromU8(files[path])) }
  const workbook = read("xl/workbook.xml"), rels = read("xl/_rels/workbook.xml.rels"), content = read("[Content_Types].xml")
  const relationships = new Map(elements(rels, "Relationship").map((r) => [r.getAttribute("Id"), r.getAttribute("Target") ?? ""]))
  const strings = files["xl/sharedStrings.xml"] ? elements(read("xl/sharedStrings.xml"), "si").map((s) => elements(s, "t").map((t) => t.textContent ?? "").join("")) : []
  const sheets = elements(workbook, "sheet").map((sheet) => {
    const target = relationships.get(sheet.getAttributeNS(REL, "id") ?? sheet.getAttribute("r:id"))
    if (!target) throw new Error("The workbook has a missing sheet relationship")
    const path = pathFor(target), document = read(path), rows = new Map<number, Map<string, Cell>>()
    for (const c of elements(document, "c")) {
      const ref = c.getAttribute("r") ?? "", match = ref.match(/^([A-Z]+)(\d+)$/)
      if (!match) continue
      const type = c.getAttribute("t"), raw = textAt(c, "v")
      const value = type === "s" ? strings[Number(raw)] ?? "" : type === "inlineStr" ? elements(c, "t").map((t) => t.textContent ?? "").join("") : type === "str" || type === "e" ? raw : raw !== "" && Number.isFinite(Number(raw)) ? Number(raw) : raw
      const row = Number(match[2]); if (!rows.has(row)) rows.set(row, new Map())
      rows.get(row)!.set(match[1], { ref, value, formula: textAt(c, "f"), error: type === "e" })
    }
    return { name: sheet.getAttribute("name") ?? "", path, document, rows }
  })
  const epoch = elements(workbook, "workbookPr")[0]?.getAttribute("date1904") === "1" ? Date.UTC(1904,0,1) : Date.UTC(1899,11,30)
  return { files, workbook, rels, content, sheets, epoch }
}
function requiredSheet(pkg: Package, name: string, headers: string[]): Sheet {
  const sheet = pkg.sheets.find((s) => s.name === name)
  if (!sheet) throw new Error(`Upload the updated tracker with a '${name}' sheet`)
  const actual = [...(sheet.rows.get(1)?.values() ?? [])].map((c) => String(c.value).trim())
  const missing = headers.filter((h,index) => actual[index] !== h)
  if (missing.length) throw new Error(`'${name}' does not match the updated tracker layout. Expected columns in order: ${headers.join(", ")}. No manual mapping is required for the supplied template.`)
  return sheet
}
const stringCell = (row: Map<string, Cell>, column: string) => row.get(column)?.error ? "" : String(row.get(column)?.value ?? "")
function numericCell(row: Map<string, Cell>, column: string): number {
  const cell = row.get(column)
  if (cell?.error) throw new Error(`Resolve the Excel error at ${cell.ref} before uploading`)
  if (typeof cell?.value === "number") return cell.value
  if (cell?.formula) {
    // Supports arithmetic used in receipt values and due dates; never executes Excel text as code.
    const replaced = cell.formula.replace(/\$?([A-Z]+)\$?\d+/g, (_, col: string) => String(row.get(col)?.value ?? ""))
    const result = arithmetic(replaced)
    if (result !== null) return result
    throw new Error(`Open and save the workbook in Excel to calculate ${cell.ref}, then upload it again`)
  }
  const value = Number(String(cell?.value ?? "").replace(/,/g, ""))
  if (!Number.isFinite(value)) throw new Error(`Invalid number at ${column}${cell?.ref.match(/\d+/)?.[0] ?? ""}`)
  return value
}
function arithmetic(source: string): number | null {
  if (!/^[\d.\s+\-*/()]+$/.test(source)) return null
  const tokens = source.match(/\d+(?:\.\d+)?|[()+\-*/]/g) ?? []; let cursor = 0
  const atom = (): number => { const token = tokens[cursor++]; if (token === "+") return atom(); if (token === "-") return -atom(); if (token === "(") { const n = expression(); if (tokens[cursor++] !== ")") throw new Error(); return n }; if (!token || !/^\d/.test(token)) throw new Error(); return Number(token) }
  const product = (): number => { let n = atom(); while (["*", "/"].includes(tokens[cursor])) { const op = tokens[cursor++], next = atom(); n = op === "*" ? n * next : n / next }; return n }
  const expression = (): number => { let n = product(); while (["+", "-"].includes(tokens[cursor])) { const op = tokens[cursor++], next = product(); n = op === "+" ? n + next : n - next }; return n }
  try { const n = expression(); return cursor === tokens.length && Number.isFinite(n) ? n : null } catch { return null }
}
function dateCell(row: Map<string, Cell>, column: string, epoch: number): string {
  const cell = row.get(column)
  if (!cell || cell.value === "" && !cell.formula) return ""
  if (typeof cell.value === "number" || cell.formula) {
    const serial = numericCell(row, column)
    if (serial > 0 && serial < 150000) return new Date(epoch + Math.floor(serial) * 86400000).toISOString().slice(0, 10)
  }
  const s = String(cell.value).trim()
  if (validDate(s)) return s
  const match = s.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2}|\d{4})$/)
  const iso = match ? `${match[3].length === 2 ? "20"+match[3] : match[3]}-${match[2].padStart(2,"0")}-${match[1].padStart(2,"0")}` : ""
  return validDate(iso) ? iso : ""
}
function tableRows(pkg: Package, definition: {name:string;headers:string[];required?:number}): Scalar[][] {
  const sheet = pkg.sheets.find((s) => s.name === definition.name)
  if (!sheet) return []
  const headers = [...(sheet.rows.get(1)?.values() ?? [])].map((c) => String(c.value))
  // Columns after `required` were added later; workbooks exported before then may not have them.
  if (definition.headers.some((h,i) => h !== headers[i] && (i < (definition.required ?? definition.headers.length) || (headers[i] ?? "") !== ""))) throw new Error(`The '${definition.name}' sheet has a different format. Rename it before importing so it can be preserved.`)
  return [...sheet.rows].filter(([row]) => row > 1).map(([,row]) => definition.headers.map((_,i) => row.get(columnName(i))?.value ?? "")).filter((row) => row[0] !== "")
}
export function importDebtors(bytes: Uint8Array, fileName: string): WorkbookState {
  if (!fileName.toLowerCase().endsWith(".xlsx")) throw new Error("Upload an .xlsx tracker")
  if (bytes.length > 20 * 1024 * 1024) throw new Error("Upload a workbook smaller than 20 MB")
  const pkg = openPackage(bytes)
  const debtors = requiredSheet(pkg,"Debtors Data",["Sr No","Party Name","Area","Email ID","RSE Name","Invoice No.","Invoice Date","Due Date","Invoice Amount","Amount Received","Outstanding Amount","Days Overdue","Days Outstanding from Invoice","Total OS Ageing","Overdue Ageing Bucket","Reminder Stage","Last Reminder Sent","Escalation Status","Remarks / Sales-RSE Input"])
  const receiptsSheet = requiredSheet(pkg,"Receipts",["Receipt ID","Receipt Date","Party Name","Invoice No.","Amount Received","Bank / UTR Reference","Remarks"])
  const previousSummary=pkg.sheets.find((s)=>s.name==="Desk Summary")
  if(previousSummary&&(previousSummary.rows.get(1)?.get("A")?.value!=="Measure"||previousSummary.rows.get(1)?.get("B")?.value!=="Value"))throw new Error("Rename the existing 'Desk Summary' sheet before importing; the app needs this name for its summary.")
  const invoices: WorkbookState["invoices"] = []
  for (const [r,row] of debtors.rows) {
    if (r === 1 || !stringCell(row,"B").trim() && !stringCell(row,"F").trim()) continue
    if (!stringCell(row,"B").trim() || !stringCell(row,"F").trim() || !row.get("I") || row.get("I")!.value === "" && !row.get("I")!.formula) throw new Error(`Complete the party, invoice number, and amount on Debtors Data row ${r}`)
    const party = stringCell(row,"B"), number = stringCell(row,"F")
    invoices.push({ id: invoiceId(party,number), row:r, party, number, area:stringCell(row,"C").trim(), email:stringCell(row,"D").trim(), rse:stringCell(row,"E").trim(), issued:dateCell(row,"G",pkg.epoch), due:dateCell(row,"H",pkg.epoch), amount:numericCell(row,"I"), legacyReminder:typeof row.get("Q")?.value === "number" ? dateCell(row,"Q",pkg.epoch) : stringCell(row,"Q"), remarks:stringCell(row,"S") })
  }
  if (!invoices.length) throw new Error("No invoice records were found in Debtors Data")
  if (invoices.length > 25000) throw new Error("This workspace supports up to 25,000 invoices")
  const counts = new Map<string,number>(); invoices.forEach((i) => counts.set(i.id,(counts.get(i.id) ?? 0)+1))
  invoices.forEach((i) => { if (counts.get(i.id)! > 1) i.id += `#${i.row}` })
  const receipts: Receipt[] = []
  for (const [r,row] of receiptsSheet.rows) {
    if (r === 1 || !["B","C","D","E"].some((c) => stringCell(row,c) !== "")) continue
    if (!stringCell(row,"C").trim() || !stringCell(row,"D").trim() || stringCell(row,"E") === "" && !row.get("E")?.formula) throw new Error(`Complete the party, invoice number, and amount on Receipts row ${r}`)
    const rawId = stringCell(row,"A")
    receipts.push({ id: /^[\da-f]{8}-/i.test(rawId) ? rawId : `receipt-row-${r}`, row:r, party:stringCell(row,"C"), invoice:stringCell(row,"D"), date:dateCell(row,"B",pkg.epoch), amount:numericCell(row,"E"), reference:stringCell(row,"F"), remarks:stringCell(row,"G") })
  }
  const j = (value: Scalar) => { const parsed: unknown = JSON.parse(String(value)); if (!Array.isArray(parsed) || !parsed.every((v) => typeof v === "string")) throw new Error("An invoice list in Reminder Log is invalid"); return parsed as string[] }
  const settings = { ...defaultSettings }
  for (const [key,value] of tableRows(pkg,definitions.settings)) {
    if (key === "preDueDays" || key === "repeatDays") settings[key] = Math.max(0,Math.min(365,Number(value) || 0))
    else if (["reminderSubject","reminderBody","balanceSubject","balanceBody"].includes(String(key))) settings[key as "reminderSubject"] = String(value)
  }
  const followups: WorkbookState["followups"] = {}
  for (const [id,flag,escalation,hold,note] of tableRows(pkg,definitions.followups)) followups[String(id)] = { flag:String(flag)==="1", escalation:String(escalation), hold:String(hold)==="1", note:String(note) }
  const importNotes: string[] = []
  const formula = pkg.sheets.find((s) => s.name === "Dashboard")?.rows.get(5)?.get("A")?.formula ?? ""
  const lastReferenced = Number(formula.match(/K2:K(\d+)/)?.[1])
  if (lastReferenced && lastReferenced < Math.max(...invoices.map((i)=>i.row))) importNotes.push(`The original Excel Dashboard stops at row ${lastReferenced}. This dashboard includes all ${invoices.length.toLocaleString("en-IN")} invoice rows. The original sheet is preserved on export; Desk Summary contains the full totals.`)
  if (elements(debtors.document,"f").some((f) => /\[\d+\]/.test(f.textContent ?? ""))) importNotes.push("Some legacy reminder dates use external workbook links. Imported dates are treated as historical notes; new activity is kept in the workbook logs.")
  return { version:1, fileName, importedAt:new Date().toISOString(), revision:0, invoices, receipts, settings, followups, importNotes,
    activity:tableRows(pkg,definitions.activity).map(([id,at,action,invoiceId,party,invoice,detail,before,after]) => ({id:String(id),at:String(at),action:String(action),invoiceId:String(invoiceId),party:String(party),invoice:String(invoice),detail:String(detail),before:String(before),after:String(after)})),
    reminders:tableRows(pkg,definitions.reminders).map(([id,at,date,type,party,email,ids,numbers,amount,stage,subject,body,status,sentAt,sentBy,sentTo,result]) => {
      if (!["Reminder","Balance confirmation"].includes(String(type)) || !["Prepared","Cancelled","Accepted","Failed"].includes(String(status))) throw new Error("Reminder Log has an unrecognized type or status")
      return {id:String(id),at:String(at),date:String(date),type:String(type) as "Reminder"|"Balance confirmation",party:String(party),email:String(email),invoiceIds:j(ids),invoiceNumbers:j(numbers),amount:Number(amount),stage:String(stage),subject:String(subject),body:String(body),status:String(status) as "Prepared"|"Cancelled"|"Accepted"|"Failed",...(sentAt?{sentAt:String(sentAt)}:{}),...(sentBy?{sentBy:String(sentBy)}:{}),...(sentTo?{sentTo:String(sentTo)}:{}),...(result?{result:String(result)}:{})}
    }),
    confirmations:tableRows(pkg,definitions.confirmations).map(([id,at,party,date,amount,status,reference,note]) => { if (!["Confirmed","Disputed"].includes(String(status))) throw new Error("Balance Confirmations has an unrecognized status"); return {id:String(id),at:String(at),party:String(party),date:String(date),amount:Number(amount),status:String(status) as "Confirmed"|"Disputed",reference:String(reference),note:String(note)} }),
  }
}
export function columnName(index: number): string { let name=""; for (let n=index+1;n>0;n=Math.floor((n-1)/26)) name=String.fromCharCode(65+(n-1)%26)+name; return name }
function escape(value: Scalar): string { return String(value ?? "").replace(/[&<>"']/g,(s)=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&apos;"})[s]!).replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g,"") }
function newSheet(rows: Scalar[][], headerStyle: number): Uint8Array {
  const max = Math.max(1,...rows.map((r)=>r.length))
  return strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="${MAIN}"><dimension ref="A1:${columnName(max-1)}${rows.length}"/><sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols><col min="1" max="${max}" width="24" customWidth="1"/></cols><sheetData>${rows.map((row,i)=>`<row r="${i+1}">${row.map((v,j)=>`<c r="${columnName(j)}${i+1}"${i===0 ? ` s="${headerStyle}"` : ""}${typeof v==="number" && Number.isFinite(v) ? `><v>${v}</v>` : ` t="inlineStr"><is><t xml:space="preserve">${escape(v)}</t></is>`}</c>`).join("")}</row>`).join("")}</sheetData><autoFilter ref="A1:${columnName(max-1)}${rows.length}"/></worksheet>`)
}
function setCell(sheet: Sheet, ref: string, value: Scalar, preserveFormula = false, style?: string): void {
  const rowNo = Number(ref.match(/\d+/)![0]), column = ref.match(/[A-Z]+/)![0]
  const data = sheet.sheetData ??= elements(sheet.document,"sheetData")[0]
  sheet.editRows ??= new Map([...data.children].map((r)=>[Number(r.getAttribute("r")),r]))
  sheet.editCells ??= new Map(elements(data,"c").map((c)=>[c.getAttribute("r")!,c]))
  let row = sheet.editRows.get(rowNo)
  if (!row) { row = sheet.document.createElementNS(MAIN,"row"); row.setAttribute("r",String(rowNo)); const next=[...data.children].find((r)=>Number(r.getAttribute("r"))>rowNo); data.insertBefore(row,next ?? null);sheet.editRows.set(rowNo,row) }
  let cell = sheet.editCells.get(ref)
  if (!cell) { cell=sheet.document.createElementNS(MAIN,"c"); cell.setAttribute("r",ref); const colIndex=(s:string)=>[...s].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0); const next=[...row.children].find((c)=>colIndex(c.getAttribute("r")!.match(/[A-Z]+/)![0])>colIndex(column)); row.insertBefore(cell,next ?? null); if(style)cell.setAttribute("s",style);sheet.editCells.set(ref,cell) }
  for (const child of [...cell.children]) if (child.localName!=="f" || !preserveFormula) child.remove()
  cell.removeAttribute("t")
  if (typeof value === "number" && Number.isFinite(value)) { const v=sheet.document.createElementNS(MAIN,"v");v.textContent=String(value);cell.appendChild(v) }
  else if (preserveFormula) { cell.setAttribute("t","str");const v=sheet.document.createElementNS(MAIN,"v");v.textContent=String(value ?? "");cell.appendChild(v) }
  else { cell.setAttribute("t","inlineStr");const inline=sheet.document.createElementNS(MAIN,"is"),t=sheet.document.createElementNS(MAIN,"t");t.setAttributeNS("http://www.w3.org/XML/1998/namespace","xml:space","preserve");t.textContent=String(value ?? "");inline.appendChild(t);cell.appendChild(inline) }
}
export function exportDebtors(original: Uint8Array, state: WorkbookState, reviewDate: string): Uint8Array {
  const pkg=openPackage(original), debtors=pkg.sheets.find((s)=>s.name==="Debtors Data")!, receipts=pkg.sheets.find((s)=>s.name==="Receipts")!
  const computed=calculateInvoices(state,today())
  for(const i of computed) {
    if (stringCell(debtors.rows.get(i.row)!,"D")!==i.email) setCell(debtors,`D${i.row}`,i.email)
    if (stringCell(debtors.rows.get(i.row)!,"S")!==i.remarks) setCell(debtors,`S${i.row}`,i.remarks)
    // "Last Reminder Sent" gains an entry when a live reminder is sent. A few rows hold a formula here; those are left alone.
    if (stringCell(debtors.rows.get(i.row)!,"Q")!==i.legacyReminder && !debtors.rows.get(i.row)!.get("Q")?.formula) setCell(debtors,`Q${i.row}`,i.legacyReminder)
    const bucket=(days:number)=>days===0?"Not Due":days<=30?"0-30 Days":days<=45?"31-45 Days":days<=60?"46-60 Days":days<=90?"61-90 Days":">90 Days"
    const values:Record<string,Scalar>={J:i.received,K:i.outstanding,L:i.outstanding>0&&validDate(i.due)?i.overdue:"",M:i.outstanding>0&&validDate(i.issued)?i.age:"",N:i.outstanding<=0?"Paid":bucket(i.age),O:i.outstanding<=0?"Paid":bucket(i.overdue),P:i.stage,R:i.outstanding<=0?"Closed":i.overdue>=45?"Management Escalation":i.overdue>=30?"CFO Escalation":i.overdue>=15?"Overdue - Follow-up":"Normal Follow-up"}
    for(const [col,value] of Object.entries(values)) setCell(debtors,`${col}${i.row}`,value,true)
  }
  const dateStyle = elements(receipts.document,"c").find((c)=>c.getAttribute("r")==="B2")?.getAttribute("s") ?? "0"
  let nextRow=Math.max(1,...state.receipts.filter((r)=>r.row>0).map((r)=>r.row))+1
  for(const receipt of state.receipts) {
    const r=receipt.row || nextRow++, originalRow=receipts.rows.get(r)
    const values:Record<string,Scalar>={B:validDate(receipt.date)?(Date.parse(receipt.date+"T00:00:00Z")-pkg.epoch)/86400000:receipt.date,C:receipt.party,D:receipt.invoice,E:receipt.amount,F:receipt.reference,G:receipt.remarks}
    if(!receipt.row) setCell(receipts,`A${r}`,receipt.id)
    for(const [col,value] of Object.entries(values)) if(!receipt.row || String(originalRow?.get(col)?.value ?? "")!==String(value)) setCell(receipts,`${col}${r}`,value,false,col==="B"?dateStyle:undefined)
  }
  const dimension=elements(receipts.document,"dimension")[0];if(dimension)dimension.setAttribute("ref",`A1:I${Math.max(1000,nextRow-1)}`)
  const filter=elements(receipts.document,"autoFilter")[0];if(filter)filter.setAttribute("ref",`A1:G${Math.max(2,nextRow-1)}`)
  for(const name of elements(pkg.workbook,"definedName")) if(name.textContent?.match(/^'?Receipts'?!/))name.textContent=`Receipts!$A$1:$G$${Math.max(2,nextRow-1)}`
  pkg.files[debtors.path]=serialise(debtors.document);pkg.files[receipts.path]=serialise(receipts.document)
  const headerStyle=Number(elements(debtors.document,"c").find((c)=>c.getAttribute("r")==="A1")?.getAttribute("s") ?? 0)
  const append=(name:string,rows:Scalar[][])=>{
    const existing=pkg.sheets.find((s)=>s.name===name)
    if(existing){pkg.files[existing.path]=newSheet(rows,headerStyle);return}
    const sheetId=Math.max(0,...elements(pkg.workbook,"sheet").map((s)=>Number(s.getAttribute("sheetId"))))+1
    const path=`xl/worksheets/jal-desk-${sheetId}.xml`,relId=`rIdJalDesk${sheetId}`
    const sheet=pkg.workbook.createElementNS(MAIN,"sheet");sheet.setAttribute("name",name);sheet.setAttribute("sheetId",String(sheetId));sheet.setAttributeNS(REL,"r:id",relId);elements(pkg.workbook,"sheets")[0].appendChild(sheet)
    const relation=pkg.rels.createElementNS(PKG,"Relationship");relation.setAttribute("Id",relId);relation.setAttribute("Type",`${REL}/worksheet`);relation.setAttribute("Target",`worksheets/jal-desk-${sheetId}.xml`);pkg.rels.documentElement.appendChild(relation)
    const override=pkg.content.createElementNS(CONTENT,"Override");override.setAttribute("PartName","/"+path);override.setAttribute("ContentType","application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");pkg.content.documentElement.appendChild(override)
    pkg.files[path]=newSheet(rows,headerStyle)
  }
  append(definitions.activity.name,[definitions.activity.headers,...state.activity.map((a)=>[a.id,a.at,a.action,a.invoiceId,a.party,a.invoice,a.detail,a.before,a.after])])
  append(definitions.reminders.name,[definitions.reminders.headers,...state.reminders.map((r)=>[r.id,r.at,r.date,r.type,r.party,r.email,JSON.stringify(r.invoiceIds),JSON.stringify(r.invoiceNumbers),r.amount,r.stage,r.subject,r.body,r.status,r.sentAt??"",r.sentBy??"",r.sentTo??"",r.result??""])])
  append(definitions.followups.name,[definitions.followups.headers,...Object.entries(state.followups).map(([id,f])=>[id,f.flag?1:0,f.escalation,f.hold?1:0,f.note])])
  append(definitions.confirmations.name,[definitions.confirmations.headers,...state.confirmations.map((c)=>[c.id,c.at,c.party,c.date,c.amount,c.status,c.reference,c.note])])
  append(definitions.settings.name,[definitions.settings.headers,...Object.entries(state.settings)])
  const summary=calculateInvoices(state,reviewDate),open=summary.filter((i)=>i.outstanding>0)
  append("Desk Summary",[["Measure","Value"],["Review date",reviewDate],["Workbook",state.fileName],["Invoices",summary.length],["Open invoices",open.length],["Outstanding (INR)",open.reduce((s,i)=>s+i.outstanding,0)],["Overdue (INR)",open.filter((i)=>i.overdue>0).reduce((s,i)=>s+i.outstanding,0)],["Reminder drafts",state.reminders.filter((r)=>r.status==="Prepared").length],["Note","Balances include all receipts. Review date controls ageing; this is not a historical ledger."],["Original Dashboard","Preserved as supplied. Its fixed ranges may omit invoice rows."]])
  // Excel recalculates preserved formulas when opened; no broken calculation-chain references.
  delete pkg.files["xl/calcChain.xml"]
  for(const rel of elements(pkg.rels,"Relationship"))if(rel.getAttribute("Type")?.endsWith("/calcChain"))rel.remove()
  for(const o of elements(pkg.content,"Override"))if(o.getAttribute("PartName")==="/xl/calcChain.xml")o.remove()
  let calc=elements(pkg.workbook,"calcPr")[0];if(!calc){calc=pkg.workbook.createElementNS(MAIN,"calcPr");pkg.workbook.documentElement.appendChild(calc)}calc.setAttribute("fullCalcOnLoad","1");calc.setAttribute("forceFullCalc","1");calc.setAttribute("calcMode","auto")
  // Make the new logs visible as soon as the exported workbook is opened.
  let views=elements(pkg.workbook,"bookViews")[0]
  if(!views){views=pkg.workbook.createElementNS(MAIN,"bookViews");pkg.workbook.documentElement.insertBefore(views,elements(pkg.workbook,"sheets")[0])}
  let view=elements(views,"workbookView")[0]
  if(!view){view=pkg.workbook.createElementNS(MAIN,"workbookView");views.appendChild(view)}
  const logIndex=elements(pkg.workbook,"sheet").findIndex((s)=>s.getAttribute("name")===definitions.activity.name)
  view.setAttribute("activeTab",String(logIndex));view.setAttribute("firstSheet",String(logIndex))
  pkg.files["xl/workbook.xml"]=serialise(pkg.workbook);pkg.files["xl/_rels/workbook.xml.rels"]=serialise(pkg.rels);pkg.files["[Content_Types].xml"]=serialise(pkg.content)
  return zipSync(pkg.files,{level:6})
}
