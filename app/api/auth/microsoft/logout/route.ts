import { deleteSession } from "@/lib/server/database"
import { cookie, getCookie } from "@/lib/server/microsoft"

export async function GET(request: Request) {
  const id = getCookie(request, "payment_outlook_session")
  if (id) await deleteSession(id)
  return new Response(null, {
    status: 302,
    headers: { location: new URL("/", request.url).toString(), "set-cookie": cookie("payment_outlook_session", "", 0) },
  })
}
