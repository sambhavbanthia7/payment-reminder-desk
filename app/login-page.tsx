"use client"

import { useState } from "react"
import Image from "next/image"
import { LockKeyhole, Mail } from "lucide-react"

/** Sign in with the finance mailbox. The same address and password are used in Outlook. */
export default function LoginPage({ allowed, problem }: { allowed: string[]; problem: string }) {
  const [email, setEmail] = useState(allowed.length === 1 ? allowed[0] : "")
  const [password, setPassword] = useState("")
  const [error, setError] = useState(problem)
  const [working, setWorking] = useState(false)

  const submit = async (event: React.FormEvent) => {
    event.preventDefault()
    setWorking(true); setError("")
    try {
      const response = await fetch("/api/auth/login", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ email, password }) })
      const data = await response.json().catch(() => ({})) as { error?: string }
      if (!response.ok) { setError(data.error ?? "Sign-in failed"); setWorking(false); return }
      window.location.reload()
    } catch {
      setError("The app could not be reached. Check the connection and try again."); setWorking(false)
    }
  }

  return <div className="debtors login-shell">
    <form className="login-card" onSubmit={submit}>
      <Image className="jal-logo" src="/jal-logo.png" alt="JAL" width={266} height={108} unoptimized priority/>
      <div>
        <h1>Receivables Desk</h1>
        <p>Sign in with the finance mailbox. Use the same email address and password as in Outlook. Reminders are sent from this account.</p>
      </div>
      <label className="desk-field"><span>Email address</span>
        <div className="login-input"><Mail size={16}/><input id="login-email" type="email" autoComplete="username" required value={email} onChange={(e) => setEmail(e.target.value)} placeholder={allowed[0] ?? "finance@…"}/></div>
      </label>
      <label className="desk-field"><span>Password</span>
        <div className="login-input"><LockKeyhole size={16}/><input id="login-password" type="password" autoComplete="current-password" required value={password} onChange={(e) => setPassword(e.target.value)}/></div>
      </label>
      {error && <p className="login-error" role="alert">{error}</p>}
      <button className="desk-btn primary large" disabled={working || !allowed.length}>{working ? "Checking with the mail server…" : "Sign in"}</button>
      <p className="login-note">The password is checked with JAL&apos;s mail server and kept encrypted on this app&apos;s server so reminders can be sent. It is never stored in the browser or the Excel file.</p>
    </form>
  </div>
}
