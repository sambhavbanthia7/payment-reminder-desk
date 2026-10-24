import { sessionFromRequest } from "@/lib/server/session"
import { cookie, getCookie } from "@/lib/server/http"
import { ONEDRIVE_SCOPES, oneDriveConfig, readLink, requestToken, saveLink } from "@/lib/server/onedrive"

export async function GET(request: Request) {
  // Only a signed-in desk user may link the OneDrive account.
  if (!(await sessionFromRequest(request).catch(() => null))) return Response.redirect(new URL("/", request.url), 302)
  const url = new URL(request.url), config = oneDriveConfig()
  const back = (result: string) => {
    const headers = new Headers({ location: new URL(`/?onedrive=${result}`, request.url).toString() })
    headers.append("set-cookie", cookie(request, "onedrive_oauth_state", "", 0))
    headers.append("set-cookie", cookie(request, "onedrive_oauth_verifier", "", 0))
    return new Response(null, { status: 302, headers })
  }
  const state = url.searchParams.get("state"), code = url.searchParams.get("code"), verifier = getCookie(request, "onedrive_oauth_verifier")
  if (!config || !code || !verifier || !state || state !== getCookie(request, "onedrive_oauth_state")) return back("failed")
  try {
    const token = await requestToken(config, { grant_type: "authorization_code", code, code_verifier: verifier, scope: ONEDRIVE_SCOPES, redirect_uri: new URL("/api/onedrive/callback", request.url).toString() })
    if (!token.refresh_token) throw new Error("Microsoft did not allow offline access")
    const profile = await fetch(`${config.graphUrl}/me?$select=displayName,mail,userPrincipalName`, { headers: { authorization: `Bearer ${token.access_token}` } }).then((r) => r.json() as Promise<{ displayName?: string; mail?: string; userPrincipalName?: string }>)
    const email = (profile.mail ?? profile.userPrincipalName ?? "").toLowerCase()
    // Reconnecting the same account keeps the chosen tracker; a different account starts fresh.
    const previous = await readLink()
    const keep = previous && previous.account_email === email
    await saveLink({
      access_token: token.access_token, refresh_token: token.refresh_token, expires_at: Date.now() + Number(token.expires_in ?? 3600) * 1000,
      account_email: email, account_name: profile.displayName ?? email, item_id: keep ? previous.item_id : "", item_name: keep ? previous.item_name : "",
    })
    return back("connected")
  } catch {
    return back("failed")
  }
}
