import { writeSession } from "@/lib/server/database"
import { cookie, getCookie, microsoftConfig, requestToken } from "@/lib/server/microsoft"

export async function GET(request: Request) {
  const url = new URL(request.url)
  const config = microsoftConfig()
  const state = url.searchParams.get("state")
  const code = url.searchParams.get("code")
  const expectedState = getCookie(request, "payment_oauth_state")
  const verifier = getCookie(request, "payment_oauth_verifier")
  if (!config || !state || state !== expectedState || !code || !verifier || url.searchParams.get("error")) {
    return Response.redirect(new URL("/?auth_error=invalid_callback", request.url), 302)
  }
  try {
    const redirectUri = new URL("/api/auth/microsoft/callback", request.url).toString()
    const token = await requestToken(config, {
      grant_type: "authorization_code",
      code,
      redirect_uri: redirectUri,
      code_verifier: verifier,
      scope: "openid profile offline_access User.Read Mail.Send",
    })
    const profileResponse = await fetch("https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName", {
      headers: { authorization: `Bearer ${String(token.access_token)}` },
    })
    if (!profileResponse.ok) throw new Error("Microsoft profile could not be read")
    const profile = await profileResponse.json() as { displayName?: string; mail?: string; userPrincipalName?: string }
    const id = crypto.randomUUID()
    await writeSession({
      id,
      access_token: String(token.access_token),
      refresh_token: token.refresh_token ? String(token.refresh_token) : null,
      expires_at: Date.now() + Number(token.expires_in ?? 3600) * 1000,
      account_email: profile.mail ?? profile.userPrincipalName ?? "",
      account_name: profile.displayName ?? profile.mail ?? "Outlook user",
    })
    const headers = new Headers({ location: new URL("/", request.url).toString() })
    headers.append("set-cookie", cookie("payment_outlook_session", id, 60 * 60 * 24 * 30))
    headers.append("set-cookie", cookie("payment_oauth_state", "", 0))
    headers.append("set-cookie", cookie("payment_oauth_verifier", "", 0))
    return new Response(null, { status: 302, headers })
  } catch {
    return Response.redirect(new URL("/?auth_error=token_exchange", request.url), 302)
  }
}
