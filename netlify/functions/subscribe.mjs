// Saves (or removes) one phone's reminder subscription, tied to a young man's invite token.
import { getStore } from "@netlify/blobs";
import { createHash } from "node:crypto";

const ok = (o, status = 200) => new Response(JSON.stringify(o), { status, headers: { "content-type": "application/json" } });

export default async (req) => {
  if (req.method !== "POST" && req.method !== "DELETE") return ok({ error: "method" }, 405);
  let body;
  try { body = await req.json(); } catch { return ok({ error: "bad json" }, 400); }
  const sub = body && body.subscription;
  if (!sub || typeof sub.endpoint !== "string" || !/^https:\/\//.test(sub.endpoint)) return ok({ error: "bad subscription" }, 400);
  const key = createHash("sha256").update(sub.endpoint).digest("hex").slice(0, 40);
  const store = getStore("subs");
  if (req.method === "DELETE") { await store.delete(key); return ok({ removed: true }); }
  const invite = String(body.invite || "").toUpperCase().replace(/[^A-Z0-9]/g, "").slice(0, 16);
  if (invite.length < 6) return ok({ error: "no invite" }, 400);
  await store.setJSON(key, { invite, sub, saved: new Date().toISOString() });
  return ok({ saved: true });
};

export const config = { path: "/api/subscribe" };
