// A practice OneDrive for local testing. It imitates the small part of Microsoft sign-in and
// Microsoft Graph the Receivables Desk uses, with real version tags (eTag/cTag) and If-Match.
//   npm run onedrive:practice   → http://127.0.0.1:8787, files in .onedrive-practice/
// To imitate someone editing the file in Excel, change the .xlsx in that folder.
// To imitate the file being open and locked, create "<file name>.lock" next to it.
import http from "node:http"
import fs from "node:fs"
import path from "node:path"
import crypto from "node:crypto"
import { fileURLToPath } from "node:url"

export async function startPracticeDrive({ port = 0, dir = path.resolve(".onedrive-practice"), log = () => {} } = {}) {
  const folder = path.join(dir, "Apps", "JAL Receivables Desk")
  fs.mkdirSync(folder, { recursive: true })
  const writers = new Map() // file name → { mtimeMs, by } for saves made through this server
  const idFor = (name) => Buffer.from(name).toString("base64url")
  const nameFor = (id) => Buffer.from(id, "base64url").toString("utf8")
  const item = (name) => {
    const file = path.join(folder, name)
    const stat = fs.statSync(file)
    const hash = crypto.createHash("sha1").update(fs.readFileSync(file)).digest("hex")
    const known = writers.get(name)
    const by = known && known.mtimeMs === stat.mtimeMs ? known.by : "Someone in Excel"
    return { id: idFor(name), name, size: stat.size, eTag: `"{${idFor(name)}},${Math.round(stat.mtimeMs)}"`, cTag: `"c:{${idFor(name)}},${hash}"`, lastModifiedDateTime: stat.mtime.toISOString(), lastModifiedBy: { user: { displayName: by } }, file: { mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" } }
  }
  const write = (name, bytes) => {
    fs.writeFileSync(path.join(folder, name), bytes)
    writers.set(name, { mtimeMs: fs.statSync(path.join(folder, name)).mtimeMs, by: "Practice Account" })
  }
  const send = (res, status, body, headers = {}) => {
    res.writeHead(status, { "content-type": "application/json", ...headers })
    res.end(body === undefined ? "" : JSON.stringify(body))
  }
  const error = (res, status, message) => send(res, status, { error: { code: String(status), message } })
  const body = (req) => new Promise((resolve) => { const chunks = []; req.on("data", (c) => chunks.push(c)); req.on("end", () => resolve(Buffer.concat(chunks))) })

  const server = http.createServer(async (req, res) => {
    const url = new URL(req.url, "http://127.0.0.1")
    const p = decodeURIComponent(url.pathname)
    try {
      if (req.method === "GET" && /\/oauth2\/v2\.0\/authorize$/.test(p)) {
        // A real sign-in page would appear here. The practice account signs in instantly.
        const back = new URL(url.searchParams.get("redirect_uri"))
        back.searchParams.set("code", "practice-code"); back.searchParams.set("state", url.searchParams.get("state") ?? "")
        log("[sign-in] practice account signed in")
        return send(res, 302, undefined, { location: back.toString() })
      }
      if (req.method === "POST" && /\/oauth2\/v2\.0\/token$/.test(p)) {
        const form = new URLSearchParams((await body(req)).toString())
        if (!form.get("client_id") || !form.get("client_secret")) return send(res, 400, { error: "invalid_client", error_description: "Missing client credentials" })
        return send(res, 200, { access_token: `practice-token-${crypto.randomUUID()}`, refresh_token: "practice-refresh", expires_in: 3600, token_type: "Bearer" })
      }
      if (!p.startsWith("/v1.0/")) return error(res, 404, "Not found")
      if (!/^Bearer practice-token-/.test(req.headers.authorization ?? "")) return error(res, 401, "Not signed in")
      const route = p.slice(5)
      if (req.method === "GET" && route === "/me") return send(res, 200, { displayName: "Practice Account", userPrincipalName: "practice@onedrive.example" })
      if (req.method === "GET" && route === "/me/drive/special/approot/children") {
        return send(res, 200, { value: fs.readdirSync(folder).filter((n) => !n.endsWith(".lock") && !n.startsWith(".")).map(item) })
      }
      let match = /^\/me\/drive\/special\/approot:\/(.+):\/content$/.exec(route)
      if (req.method === "PUT" && match) {
        const name = match[1]
        if (fs.existsSync(path.join(folder, name)) && url.searchParams.get("@microsoft.graph.conflictBehavior") === "fail") return error(res, 409, "A file with this name already exists")
        write(name, await body(req)); log(`[drive] created ${name}`)
        return send(res, 201, item(name))
      }
      match = /^\/me\/drive\/items\/([^/]+)(\/content)?$/.exec(route)
      if (match) {
        const name = nameFor(match[1]), file = path.join(folder, name)
        if (!fs.existsSync(file)) return error(res, 404, "Item not found")
        if (req.method === "GET" && !match[2]) return send(res, 200, item(name))
        if (req.method === "GET") { res.writeHead(200, { "content-type": "application/octet-stream" }); return res.end(fs.readFileSync(file)) }
        if (req.method === "PUT") {
          if (fs.existsSync(`${file}.lock`)) return error(res, 423, "The resource you are attempting to access is locked")
          const ifMatch = req.headers["if-match"]
          if (ifMatch && ifMatch !== item(name).eTag) { log(`[drive] save refused: ${name} changed since it was opened`); return error(res, 412, "ETag does not match current item's value") }
          write(name, await body(req)); log(`[drive] saved ${name}`)
          return send(res, 200, item(name))
        }
      }
      return error(res, 404, "Not found")
    } catch (e) {
      return error(res, 500, e instanceof Error ? e.message : "Practice OneDrive error")
    }
  })
  await new Promise((resolve) => server.listen(port, "127.0.0.1", resolve))
  return { port: server.address().port, folder, close: () => new Promise((r) => server.close(r)) }
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const drive = await startPracticeDrive({ port: 8787, log: (line) => console.log(line) })
  console.log(`Practice OneDrive running at http://127.0.0.1:${drive.port} · files in ${drive.folder}`)
}
