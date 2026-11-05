# fsydaily.com

The website for FSY Daily Prep. It shows the Google Apps Script app full-screen, carries the home-screen icon,
and sends reminder notifications.

- `site/` is the website (Netlify publishes this folder).
- `netlify/functions/subscribe.mjs` saves a phone's reminder sign-up.
- `netlify/functions/remind.mjs` runs every hour and nudges only the young men who haven't checked in.

Netlify environment variables (Site configuration > Environment variables):
`VAPID_PUBLIC_KEY`, `VAPID_PRIVATE_KEY`, `PUSH_SECRET` (same as PUSH_SECRET in Code.gs), `APPS_SCRIPT_URL` (the /exec link).

Not an official site of The Church of Jesus Christ of Latter-day Saints.
