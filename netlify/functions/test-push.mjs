// Sends one test reminder to the phone that asked, so a young man (or leader) can see reminders work.
import webpush from "web-push";

const ok = (o, status = 200) => new Response(JSON.stringify(o), { status, headers: { "content-type": "application/json" } });

export default async (req) => {
  if (req.method !== "POST") return ok({ error: "method" }, 405);
  const { VAPID_PUBLIC_KEY, VAPID_PRIVATE_KEY } = process.env;
  if (!VAPID_PUBLIC_KEY || !VAPID_PRIVATE_KEY) return ok({ error: "Reminders aren't set up on the server yet." }, 500);
  let body;
  try { body = await req.json(); } catch { return ok({ error: "bad json" }, 400); }
  const sub = body && body.subscription;
  if (!sub || typeof sub.endpoint !== "string" || !/^https:\/\//.test(sub.endpoint)) return ok({ error: "bad subscription" }, 400);
  webpush.setVapidDetails("https://fsydaily.com", VAPID_PUBLIC_KEY, VAPID_PRIVATE_KEY);
  try {
    await webpush.sendNotification(sub, JSON.stringify({ title: "Reminders are working", body: "You'll get these at 12, 4 and 8 on days you haven't checked in.", url: "/" }), { TTL: 600 });
    return ok({ sent: true });
  } catch (e) {
    console.error("test push failed", e.statusCode || e.message, e.body || "");
    return ok({ error: "Couldn't send (" + (e.statusCode || "error") + ")." }, 502);
  }
};

export const config = { path: "/api/test-push" };
