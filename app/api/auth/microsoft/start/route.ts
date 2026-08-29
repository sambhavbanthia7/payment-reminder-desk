import { cookie, microsoftConfig, pkceChallenge, randomUrlSafe } from "@/lib/server/microsoft"

export async function GET(request: Request) {
  const config = microsoftConfig()
  if (!config) return Response.json({ error: "Outlook authentication is not configured" }, { status: 503 })
  const state = randomUrlSafe(24)
  const verifier = randomUrlSafe(48)
  const challenge = await pkceChallenge(verifier)
  const redirectUri = new URL("/api/auth/microsoft/callback", request.url).toString()
  const authorize = new URL(`https://login.microsoftonline.com/${encodeURIComponent(config.tenantId)}/oauth2/v2.0/authorize`)
  authorize.search = new URLSearchParams({
    client_id: config.clientId,
    response_type: "code",
    redirect_uri: redirectUri,
    response_mode: "query",
    scope: "openid profile offline_access User.Read Mail.Send",
    state,
    code_challenge: challenge,
    code_challenge_method: "S256",
  }).toString()
  const headers = new Headers({ location: authorize.toString() })
  headers.append("set-cookie", cookie("payment_oauth_state", state, 600))
  headers.append("set-cookie", cookie("payment_oauth_verifier", verifier, 600))
  return new Response(null, { status: 302, headers })
}
