import { requireSession } from "@/lib/server/session"
import { cookie, pkceChallenge, randomUrlSafe } from "@/lib/server/http"
import { ONEDRIVE_SCOPES, oneDriveConfig } from "@/lib/server/onedrive"

export async function GET(request: Request) {
  const signedIn = await requireSession(request)
  if (signedIn instanceof Response) return signedIn
  const config = oneDriveConfig()
  if (!config) return Response.json({ error: "OneDrive is not set up yet" }, { status: 503 })
  const state = randomUrlSafe(24), verifier = randomUrlSafe(48)
  const authorize = new URL(`${config.loginUrl}/${encodeURIComponent(config.tenant)}/oauth2/v2.0/authorize`)
  authorize.search = new URLSearchParams({
    client_id: config.clientId, response_type: "code", response_mode: "query", scope: ONEDRIVE_SCOPES, state,
    redirect_uri: new URL("/api/onedrive/callback", request.url).toString(),
    code_challenge: await pkceChallenge(verifier), code_challenge_method: "S256", prompt: "select_account",
  }).toString()
  const headers = new Headers({ location: authorize.toString() })
  headers.append("set-cookie", cookie(request, "onedrive_oauth_state", state))
  headers.append("set-cookie", cookie(request, "onedrive_oauth_verifier", verifier))
  return new Response(null, { status: 302, headers })
}
