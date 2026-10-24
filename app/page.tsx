import { headers } from "next/headers"
import DebtorsDashboard from "./debtors-dashboard"
import LoginPage from "./login-page"
import { mailSettings } from "@/lib/server/mail"
import { sessionFromCookieHeader } from "@/lib/server/session"

export const dynamic = "force-dynamic"

export default async function Home() {
  const { settings, problem } = mailSettings()
  let session = null, sessionProblem = ""
  try {
    session = await sessionFromCookieHeader((await headers()).get("cookie") ?? "")
  } catch (error) {
    sessionProblem = error instanceof Error ? error.message : "Sign-in is unavailable"
  }
  if (!session) return <LoginPage allowed={settings?.allowed ?? []} problem={sessionProblem || problem || (settings ? "" : "Email sign-in is not set up on this server yet.")}/>
  return <DebtorsDashboard account={session.email}/>
}
