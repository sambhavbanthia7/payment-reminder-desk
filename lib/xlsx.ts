import type { SheetTable } from "./reminders"

type ZipEntry = { method: number; compressedSize: number; localOffset: number }

export async function readWorkbook(file: File): Promise<SheetTable[]> {
  if (!file.name.toLowerCase().endsWith(".xlsx")) {
    throw new Error("Please upload an .xlsx file. Legacy .xls files are not supported.")
  }
  const bytes = new Uint8Array(await file.arrayBuffer())
  const entries = readZipDirectory(bytes)
  const readText = async (name: string) => {
    const entry = entries.get(name)
    if (!entry) return null
    const data = await extractZipEntry(bytes, entry)
    return new TextDecoder().decode(data)
  }

  const workbookXml = await readText("xl/workbook.xml")
  const relsXml = await readText("xl/_rels/workbook.xml.rels")
  if (!workbookXml || !relsXml) throw new Error("This workbook is missing its sheet directory.")

  const sharedStrings = await parseSharedStrings(await readText("xl/sharedStrings.xml"))
  const relationships = parseRelationships(relsXml)
  const workbookDoc = xml(workbookXml)
  const sheets = [...workbookDoc.getElementsByTagNameNS("*", "sheet")]
  const output: SheetTable[] = []

  for (const sheet of sheets) {
    const relationshipId = sheet.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id")
      ?? sheet.getAttribute("r:id")
    const target = relationshipId ? relationships.get(relationshipId) : null
    if (!target) continue
    const sheetPath = target.startsWith("/") ? target.slice(1) : `xl/${target.replace(/^\.\//, "")}`
    const sheetXml = await readText(normalisePath(sheetPath))
    if (!sheetXml) continue
    const matrix = parseWorksheet(sheetXml, sharedStrings)
    const detected = detectHeaderRow(matrix)
    if (!detected) continue
    const { rowIndex, headers } = detected
    const rows = matrix.slice(rowIndex + 1)
      .filter((row) => row.some((value) => value !== null && String(value).trim() !== ""))
      .map((row) => Object.fromEntries(headers.map((header, index) => [header, row[index] ?? null])))
    output.push({ fileName: file.name, sheetName: sheet.getAttribute("name") ?? "Sheet", headers, rows })
  }

  if (!output.length) throw new Error("No tabular sheet with a recognizable header row was found.")
  return output
}

export function selectBestSheet(sheets: SheetTable[], kind: "dues" | "master"): SheetTable {
  const desired = kind === "dues"
    ? ["accountname", "accode", "creditdays", "dated", "vrno", "vrtype", "bal"]
    : ["partycode", "name", "email"]
  return [...sheets].sort((a, b) => score(b.headers, desired) - score(a.headers, desired))[0]
}

function readZipDirectory(bytes: Uint8Array): Map<string, ZipEntry> {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  let eocd = -1
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65_557); i--) {
    if (view.getUint32(i, true) === 0x06054b50) { eocd = i; break }
  }
  if (eocd < 0) throw new Error("The uploaded file is not a valid XLSX archive.")
  const entryCount = view.getUint16(eocd + 10, true)
  let cursor = view.getUint32(eocd + 16, true)
  const entries = new Map<string, ZipEntry>()
  const decoder = new TextDecoder()
  for (let i = 0; i < entryCount; i++) {
    if (view.getUint32(cursor, true) !== 0x02014b50) break
    const method = view.getUint16(cursor + 10, true)
    const compressedSize = view.getUint32(cursor + 20, true)
    const fileNameLength = view.getUint16(cursor + 28, true)
    const extraLength = view.getUint16(cursor + 30, true)
    const commentLength = view.getUint16(cursor + 32, true)
    const localOffset = view.getUint32(cursor + 42, true)
    const name = decoder.decode(bytes.slice(cursor + 46, cursor + 46 + fileNameLength))
    entries.set(name, { method, compressedSize, localOffset })
    cursor += 46 + fileNameLength + extraLength + commentLength
  }
  return entries
}

async function extractZipEntry(bytes: Uint8Array, entry: ZipEntry): Promise<Uint8Array> {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const cursor = entry.localOffset
  if (view.getUint32(cursor, true) !== 0x04034b50) throw new Error("The workbook contains a damaged ZIP entry.")
  const fileNameLength = view.getUint16(cursor + 26, true)
  const extraLength = view.getUint16(cursor + 28, true)
  const start = cursor + 30 + fileNameLength + extraLength
  const compressed = bytes.slice(start, start + entry.compressedSize)
  if (entry.method === 0) return compressed
  if (entry.method !== 8 || typeof DecompressionStream === "undefined") {
    throw new Error("This browser cannot decompress the uploaded workbook.")
  }
  const stream = new Blob([compressed]).stream().pipeThrough(new DecompressionStream("deflate-raw"))
  return new Uint8Array(await new Response(stream).arrayBuffer())
}

async function parseSharedStrings(source: string | null): Promise<string[]> {
  if (!source) return []
  const document = xml(source)
  return [...document.getElementsByTagNameNS("*", "si")].map((item) =>
    [...item.getElementsByTagNameNS("*", "t")].map((node) => node.textContent ?? "").join(""),
  )
}

function parseRelationships(source: string): Map<string, string> {
  const document = xml(source)
  return new Map([...document.getElementsByTagNameNS("*", "Relationship")].map((rel) => [rel.getAttribute("Id") ?? "", rel.getAttribute("Target") ?? ""]))
}

function parseWorksheet(source: string, sharedStrings: string[]): unknown[][] {
  const document = xml(source)
  const matrix: unknown[][] = []
  for (const cell of [...document.getElementsByTagNameNS("*", "c")]) {
    const ref = cell.getAttribute("r") ?? "A1"
    const match = ref.match(/^([A-Z]+)(\d+)$/)
    if (!match) continue
    const row = Number(match[2]) - 1
    const col = columnNumber(match[1])
    matrix[row] ??= []
    const type = cell.getAttribute("t")
    const valueNode = cell.getElementsByTagNameNS("*", "v")[0]
    let value: unknown = valueNode?.textContent ?? null
    if (type === "s") value = sharedStrings[Number(value)] ?? ""
    else if (type === "inlineStr") value = [...cell.getElementsByTagNameNS("*", "t")].map((node) => node.textContent ?? "").join("")
    else if (type === "b") value = value === "1"
    else if (type !== "str" && value !== null && value !== "") {
      const number = Number(value)
      if (Number.isFinite(number)) value = number
    }
    matrix[row][col] = value
  }
  return matrix
}

function detectHeaderRow(matrix: unknown[][]): { rowIndex: number; headers: string[] } | null {
  let best: { rowIndex: number; headers: string[]; score: number } | null = null
  matrix.slice(0, 20).forEach((row, rowIndex) => {
    const values = row.map((value) => String(value ?? "").trim())
    const populated = values.filter(Boolean).length
    const unique = new Set(values.filter(Boolean).map((value) => value.toLowerCase())).size
    const headerLike = values.filter((value) => /[a-z]/i.test(value)).length
    const rowScore = populated + unique + headerLike * 2
    if (populated >= 3 && (!best || rowScore > best.score)) {
      const seen = new Map<string, number>()
      const headers = values.map((value, index) => {
        const base = value || `Column ${index + 1}`
        const count = seen.get(base) ?? 0
        seen.set(base, count + 1)
        return count ? `${base} (${count + 1})` : base
      })
      best = { rowIndex, headers, score: rowScore }
    }
  })
  return best
}

function xml(source: string): XMLDocument {
  const document = new DOMParser().parseFromString(source, "application/xml")
  if (document.getElementsByTagName("parsererror").length) throw new Error("The workbook contains invalid XML.")
  return document
}

function columnNumber(letters: string): number {
  return [...letters].reduce((value, letter) => value * 26 + letter.charCodeAt(0) - 64, 0) - 1
}

function normalisePath(path: string): string {
  const parts: string[] = []
  for (const part of path.split("/")) {
    if (part === "..") parts.pop()
    else if (part !== "." && part) parts.push(part)
  }
  return parts.join("/")
}

function score(headers: string[], desired: string[]): number {
  const normalized = headers.map((header) => header.toLowerCase().replace(/[^a-z0-9]/g, ""))
  return desired.filter((field) => normalized.includes(field)).length
}
