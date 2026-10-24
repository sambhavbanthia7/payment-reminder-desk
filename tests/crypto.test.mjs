import assert from "node:assert/strict"
import test from "node:test"
import { seal, unseal } from "../lib/server/crypto.ts"

test("a sealed password opens only with the same secret and is not stored in plain text", async () => {
  const sealed = await seal("s".repeat(40), "[pass%word?")
  assert.doesNotMatch(sealed, /pass%word/)
  assert.equal(await unseal("s".repeat(40), sealed), "[pass%word?")
  assert.equal(await unseal("t".repeat(40), sealed), null)
  assert.notEqual(await seal("s".repeat(40), "same"), await seal("s".repeat(40), "same"))
})

test("a tampered value is rejected", async () => {
  const sealed = await seal("k".repeat(40), "secret")
  const [iv, data] = sealed.split(".")
  const flipped = data.slice(0, 4) + (data[4] === "A" ? "B" : "A") + data.slice(5)
  assert.equal(await unseal("k".repeat(40), `${iv}.${flipped}`), null)
})
