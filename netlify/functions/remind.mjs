// Runs every hour. Asks the app which young men haven't checked in yet and nudges only them.
import { getStore } from "@netlify/blobs";
import webpush from "web-push";

const ALL = [0, 1, 2, 3, 4, 5, 6];
const MESSAGES = {
  noon: { hour: 12, days: ALL, title: "Today's FSY reading is waiting", body: "Two minutes. Read it, think about it, check in." },
  afternoon: { hour: 16, days: ALL, title: "Still time for today's check-in", body: "Read today's scripture and share one thought." },
  saturdayNight: { hour: 20, days: [6], title: "Tomorrow's lesson: bring one thing", body: "Do tonight's reading so you're ready for Sunday." },
  evening: { hour: 20, days: [0, 1, 2, 3, 4, 5], title: "Last call for today's check-in", body: "Still time to read today's scripture before bed." }
};

function pick(hour, dow) {
  for (const m of Object.values(MESSAGES)) if (m.hour === hour && m.days.includes(dow)) return m;
  return null;
}

export default async () => {
  const { VAPID_PUBLIC_KEY, VAPID_PRIVATE_KEY, APPS_SCRIPT_URL, PUSH_SECRET } = process.env;
  if (!VAPID_PUBLIC_KEY || !VAPID_PRIVATE_KEY || !APPS_SCRIPT_URL || !PUSH_SECRET) {
    console.error("Missing settings: VAPID_PUBLIC_KEY, VAPID_PRIVATE_KEY, APPS_SCRIPT_URL, PUSH_SECRET");
    return;
  }
  webpush.setVapidDetails("https://fsydaily.com", VAPID_PUBLIC_KEY, VAPID_PRIVATE_KEY);

  const store = getStore("subs");
  const { blobs } = await store.list();
  if (!blobs.length) return;
  const subs = [];
  for (const b of blobs) {
    const v = await store.get(b.key, { type: "json" });
    if (v && v.invite && v.sub) subs.push({ key: b.key, ...v });
  }
  const tokens = [...new Set(subs.map((s) => s.invite))];

  const res = await fetch(APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ secret: PUSH_SECRET, action: "status", tokens }),
    redirect: "follow"
  });
  const data = await res.json().catch(() => ({}));
  const status = data.status || {};
  if (data.error) { console.error("App said:", data.error); return; }

  let sent = 0, removed = 0;
  for (const s of subs) {
    const st = status[s.invite];
    if (!st) continue;
    if (!st.ok) { await store.delete(s.key); removed++; continue; }   // invite was reset or the name was removed
    if (st.done) continue;
    const msg = pick(st.hour, st.dow);
    if (!msg) continue;
    try {
      await webpush.sendNotification(s.sub, JSON.stringify({ title: msg.title, body: msg.body, url: "/" }), { TTL: 3600 });
      sent++;
    } catch (e) {
      if (e.statusCode === 404 || e.statusCode === 410) { await store.delete(s.key); removed++; }
      else console.error("push failed", e.statusCode || e.message);
    }
  }
  console.log(`reminders: ${sent} sent, ${removed} removed, ${subs.length} phones`);
};

export const config = { schedule: "@hourly" };
