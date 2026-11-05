// FSY_DAILY_PREP_CODE
// FSY Daily Prep: one app for many wards, backed by the Google Sheet this script is attached to.
// Tabs: Wards (ID | Name | Join code | Old leader PIN | Created | Time zone), Roster (Name | Notes | Ward ID | Invite | Joined),
//       Readings, Check-ins (Timestamp | Date | Name | Thought | Ward ID), Board.
// Greens Ferry is the default ward: the plain app link opens it, exactly like before.
// Not an official app of The Church of Jesus Christ of Latter-day Saints.

const TZ = 'America/Los_Angeles';            // used when a ward has no time zone set
const ICON_URL = 'https://cdn.prod.website-files.com/571e75edbd5a690c420161f4/6ad91704cdda56e5dd879be5_fsy-icon-1024.png';
const DEFAULT_WARD_ID = 'w_greensferry';    // rows with a blank Ward ID belong to this ward
// Private settings (setup PIN, push secret, website address) live in the sheet's Settings tab, not here,
// so this code has nothing private in it.
const SUMMARY_HOUR = 18;                     // leaders' daily email goes out at 6 PM ward time
const LEGACY_UNTIL = '2026-11-22';           // until this date, phones already set up on the old link still open Greens Ferry
const CODE_CHARS = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
const MAX_WARDS = 300;
const MAX_NAMES = 60;
const RECENT_ROWS = 15000;                   // the youth screen only scans this many latest check-ins
const ZONES = ['America/Los_Angeles', 'America/Boise', 'America/Denver', 'America/Phoenix',
               'America/Chicago', 'America/New_York', 'America/Anchorage', 'Pacific/Honolulu'];

function doGet(e) {
  const p = (e && e.parameter) || {};
  const t = HtmlService.createTemplateFromFile('Index');
  t.wardParam = norm_(p.w).slice(0, 24);
  t.nameParam = String(p.n || '').slice(0, 40);
  t.inviteParam = norm_(p.i).slice(0, 16);
  t.leaderParam = norm_(p.L).slice(0, 16);
  t.joinParam = norm_(p.j).slice(0, 24);
  t.appUrl = ScriptApp.getService().getUrl();
  t.zones = ZONES;
  return t.evaluate()
    .setTitle('FSY Daily Prep')
    .setFaviconUrl(ICON_URL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);   // lets fsydaily.com show the app
}

/* ---------- settings (Settings tab: Setting | Value) ---------- */
function settings_() {
  return cached_('settings', 300, function () {
    const o = {}, sh = sh_('Settings');
    if (sh) sh.getDataRange().getValues().slice(1).forEach(function (r) { o[String(r[0]).trim()] = String(r[1]).trim(); });
    return o;
  });
}
function setupCode_() { return settings_()['Setup PIN'] || ''; }
function pushSecret_() { return settings_()['Push secret'] || ''; }
function publicUrl_() { return (settings_()['Website address'] || '').replace(/\/+$/, ''); }

/* ---------- helpers ---------- */
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sh_(name) { return ss_().getSheetByName(name); }
function cache_() { return CacheService.getScriptCache(); }
function cached_(key, secs, fn) {
  const c = cache_(), hit = c.get(key);
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  const v = fn();
  try { const s = JSON.stringify(v); if (s.length < 95000) c.put(key, s, secs); } catch (e) {}
  return v;
}
function bust_(keys) { cache_().removeAll(keys); }
function tzOf_(w) { return (w && w.tz) || TZ; }
function today_(w) { return Utilities.formatDate(new Date(), tzOf_(w), 'yyyy-MM-dd'); }
function day_(v) { return (v instanceof Date) ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v).trim(); }
function norm_(c) { return String(c || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); }
function clean_(s, max) {
  s = String(s || '').replace(/\s+/g, ' ').trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;          // keep it as plain text in the sheet
}
function wardOf_(v) { const id = String(v || '').trim(); return id || DEFAULT_WARD_ID; }

function wards_() {
  return cached_('wards', 120, function () {
    return sh_('Wards').getDataRange().getValues().slice(1)
      .map(function (r, i) {
        return { id: String(r[0]).trim(), name: String(r[1]).trim(), youth: norm_(r[2]), youthPretty: String(r[2]).trim(), leader: norm_(r[3]),
                 tz: ZONES.indexOf(String(r[5] || '').trim()) >= 0 ? String(r[5]).trim() : '', row: i + 2 };
      })
      .filter(function (w) { return w.id; });
  });
}
function wardById_(id) { return wards_().filter(function (w) { return w.id === id; })[0] || null; }
function wardByYouth_(code, legacy) {
  code = norm_(code);
  if (!code) {
    if (legacy && Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd') <= LEGACY_UNTIL) return wardById_(DEFAULT_WARD_ID);
    throw new Error('Open the invite link your leader sent you, or enter your ward code.');
  }
  const w = wards_().filter(function (x) { return x.youth === code; })[0];
  if (!w) throw new Error("That ward code didn't work. Check it with your leader and try again.");
  return w;
}
// Leaders tab: Name | Ward ID | Link token. Each leader gets a personal link that opens the leader view.
function leadersSheet_() {
  let sh = sh_('Leaders');
  if (!sh) { sh = ss_().insertSheet('Leaders'); sh.appendRow(['Name', 'Ward ID', 'Link token', 'Added', 'Email']); sh.setFrozenRows(1); }
  return sh;
}
function leaders_() {
  return cached_('leaders', 300, function () {
    const m = {};
    leadersSheet_().getDataRange().getValues().slice(1).forEach(function (r) {
      const t = norm_(r[2]);
      if (t && String(r[0]).trim()) m[t] = { ward: String(r[1]).trim(), name: String(r[0]).trim(), email: String(r[4] || '').trim() };
    });
    return m;
  });
}
// Leaders sign in with "WARDCODE:PIN" (ward code + 4-digit PIN), a personal leader-link token,
// or an older 8-character leader code. Five wrong tries locks that ward's sign-in for 15 minutes.
function wardByLeader_(cred) {
  cred = String(cred || '');
  const parts = cred.split(':');
  let w = null, lockKey = 'fail_any';
  if (parts.length === 2) {
    const wardCode = norm_(parts[0]), pin = norm_(parts[1]);
    lockKey = 'fail_' + wardCode;
    tooMany_(lockKey);
    const ward = wards_().filter(function (x) { return x.youth === wardCode; })[0];
    if (ward && pin && ward.leader === pin) w = ward;
  } else {
    const code = norm_(cred);
    if (code.length >= 8) {
      w = wards_().filter(function (x) { return x.leader === code; })[0] || null;
      if (!w) { const hit = leaders_()[code]; if (hit) w = wardById_(hit.ward); }
    }
  }
  if (!w) {
    failed_(lockKey);
    throw new Error("That ward code and PIN didn't match. Check them and try again.");
  }
  return w;
}
function tooMany_(key) {
  if (Number(cache_().get(key) || 0) >= 5) throw new Error('Too many wrong tries. Wait 15 minutes and try again.');
}
function failed_(key) {
  const c = cache_(), n = Number(c.get(key) || 0) + 1;
  c.put(key, String(n), 900);
}
function pin_(p) {
  p = String(p || '').replace(/\D/g, '');
  if (p.length !== 4) throw new Error('Choose a PIN that is exactly 4 numbers.');
  return p;
}
function newCode_(len) {
  const used = {};
  wards_().forEach(function (w) { used[w.youth] = 1; used[w.leader] = 1; });
  let c;
  do {
    c = '';
    for (let i = 0; i < len; i++) c += CODE_CHARS.charAt(Math.floor(Math.random() * CODE_CHARS.length));
  } while (used[c]);
  return c;
}
function base_() { return publicUrl_() || ScriptApp.getService().getUrl(); }
// Youth codes read like the ward: "Timberlane Ward" -> TIMBERLANE-482
function youthCode_(wardName) {
  let base = norm_(String(wardName).replace(/\bward\b/ig, '')).slice(0, 14) || 'WARD';
  const used = {};
  wards_().forEach(function (w) { used[w.youth] = 1; used[w.leader] = 1; });
  let c;
  do { c = base + '-' + (100 + Math.floor(Math.random() * 900)); } while (used[norm_(c)]);
  return c;
}
function link_(w) { return base_() + '?w=' + norm_(w.youth); }
function inviteLink_(token) { return base_() + '?i=' + token; }
function newToken_() {
  const used = invites_();
  let c;
  do {
    c = '';
    for (let i = 0; i < 10; i++) c += CODE_CHARS.charAt(Math.floor(Math.random() * CODE_CHARS.length));
  } while (used[c]);
  return c;
}
// token -> { ward, name }
function invites_() {
  return cached_('invites', 300, function () {
    const m = {};
    sh_('Roster').getDataRange().getValues().slice(1).forEach(function (r) {
      const t = norm_(r[3]);
      if (t && String(r[0]).trim()) m[t] = { ward: wardOf_(r[2]), name: String(r[0]).trim(), joined: !!String(r[4] || '').trim() };
    });
    return m;
  });
}
function byInvite_(token) {
  const hit = invites_()[norm_(token)];
  const w = hit ? wardById_(hit.ward) : null;
  if (!w) throw new Error("Your old sign-in doesn't work anymore. Ask your leader for your quorum's new join link, open it, and tap your name.");
  if (!hit.joined) markJoined_(w.id, hit.name);
  return { ward: w, name: hit.name };
}
// Roster rows for one ward, straight from the sheet: { row, name, token, joined }
function rosterRows_(wardId) {
  const out = [];
  sh_('Roster').getDataRange().getValues().forEach(function (r, i) {
    if (i === 0 || !String(r[0]).trim() || wardOf_(r[2]) !== wardId) return;
    out.push({ row: i + 1, name: String(r[0]).trim(), token: norm_(r[3]), joined: String(r[4] || '').trim() });
  });
  return out;
}
function markJoined_(wardId, name) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    rosterRows_(wardId).forEach(function (r) {
      if (r.name === name && !r.joined) sh_('Roster').getRange(r.row, 5).setValue("'" + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'));
    });
    bust_(['invites']);
  } finally { lock.releaseLock(); }
}
function wardByJoin_(code) {
  code = norm_(code);
  const w = code ? wards_().filter(function (x) { return x.youth === code; })[0] : null;
  if (!w) throw new Error("That join link doesn't work anymore. Ask your leader for the current one.");
  return w;
}
function joinLink_(w) { return base_() + '?j=' + norm_(w.youth); }
function sendLeaderLink_(to, wardName, token, intro) {
  const link = base_() + '?L=' + token;
  const name = String(wardName).replace(/^'/, '').replace(/[<>&]/g, '');
  MailApp.sendEmail({ to: to, name: 'FSY Daily Prep', subject: 'Your FSY Daily Prep sign-in link: ' + name,
    htmlBody: '<div style="font-family:Arial,sans-serif;font-size:15px;color:#18212b">' +
      '<p>' + intro + '</p>' +
      '<p><a href="' + link + '" style="display:inline-block;padding:12px 18px;background:#b5821f;color:#1a1205;border-radius:8px;text-decoration:none;font-weight:bold">Open the ' + name + ' dashboard</a></p>' +
      '<p style="color:#5b6774;font-size:13px">This link signs you in, so keep it to yourself. Your phone stays signed in after you open it. Save it to your home screen for quick access.</p></div>' });
}
// gives every name in a ward an invite token (run under the script lock)
function ensureTokens_(wardId) {
  const sh = sh_('Roster'), vals = sh.getDataRange().getValues();
  let changed = false;
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() && wardOf_(vals[i][2]) === wardId && !norm_(vals[i][3])) {
      const t = newToken_();
      sh.getRange(i + 1, 4).setValue(t);
      const m = invites_(); m[t] = { ward: wardId, name: String(vals[i][0]).trim() };
      cache_().put('invites', JSON.stringify(m), 300);
      changed = true;
    }
  }
  if (changed) bust_(['invites']);
}
function inviteList_(wardId) {
  const out = [];
  sh_('Roster').getDataRange().getValues().slice(1).forEach(function (r) {
    if (String(r[0]).trim() && wardOf_(r[2]) === wardId && norm_(r[3])) out.push({ name: String(r[0]).trim(), link: inviteLink_(norm_(r[3])) });
  });
  return out;
}

function roster_(wardId) {
  return cached_('roster_' + wardId, 300, function () {
    return sh_('Roster').getDataRange().getValues().slice(1)
      .filter(function (r) { return String(r[0]).trim() && wardOf_(r[2]) === wardId; })
      .map(function (r) { return String(r[0]).trim(); });
  });
}
function rows_(recentOnly) {
  const sh = sh_('Check-ins'), last = sh.getLastRow();
  if (last < 2) return [];
  const start = recentOnly ? Math.max(2, last - RECENT_ROWS + 1) : 2;
  return sh.getRange(start, 1, last - start + 1, 5).getValues();
}
function checkins_(wardId) {
  return cached_('ck_' + wardId, 300, function () {
    return rows_(true)
      .filter(function (r) { return String(r[2]).trim() && wardOf_(r[4]) === wardId; })
      .map(function (r) { return { date: day_(r[1]), name: String(r[2]).trim() }; });
  });
}
function checkinsFull_(wardId) {
  return rows_(false)
    .filter(function (r) { return String(r[2]).trim() && wardOf_(r[4]) === wardId; })
    .map(function (r) {
      return { date: day_(r[1]), name: String(r[2]).trim(), thought: String(r[3] || ''),
               time: (r[0] instanceof Date) ? Utilities.formatDate(r[0], TZ, 'EEE M/d h:mm a') : String(r[0]) };
    });
}
function readings_() {
  return cached_('readings', 600, function () {
    return sh_('Readings').getDataRange().getValues().slice(1)
      .filter(function (r) { return r[0] !== ''; })
      .map(function (r) {
        return { date: day_(r[0]), title: r[1], ref: r[2], refUrl: r[3], thought: r[4],
                 ask: r[5], lesson: r[6], lessonUrl: r[7] };
      });
  });
}

/* ---------- youth ---------- */
// Boys: list of names for a ward's join link (no answers, no tokens).
function joinInfo(code) {
  const w = wardByJoin_(code);
  return { wardName: String(w.name).replace(/^'/, ''), names: rosterRows_(w.id).map(function (r) { return { name: r.name, joined: !!r.joined }; }) };
}
// A boy taps his name on the join page. The name locks to his phone until a leader taps Reset.
// Phones already set up the old way (legacy) can keep their name without a reset.
function claimName(code, name, legacy, leaderCred) {
  name = String(name || '').trim();
  let w;
  if (code) w = wardByJoin_(code);
  else w = wardByYouth_('', legacy);
  // A signed-in leader of this ward can open his own reading on any phone without a reset.
  let isLeader = false;
  if (leaderCred) { try { isLeader = wardByLeader_(leaderCred).id === w.id; } catch (x) {} }
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const r = rosterRows_(w.id).filter(function (x) { return x.name === name; })[0];
    if (!r) throw new Error("That name isn't on your quorum's list. Ask your leader to add you.");
    const oldPhone = legacy && Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd') <= LEGACY_UNTIL;
    if (r.joined && r.token && !oldPhone && !isLeader) throw new Error('TAKEN');
    let token = r.token;
    if (!token) { token = newToken_(); sh_('Roster').getRange(r.row, 4).setValue(token); }
    if (!r.joined) sh_('Roster').getRange(r.row, 5).setValue("'" + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'));
    bust_(['invites']);
    return { invite: token };
  } finally {
    lock.releaseLock();
  }
}
// A boy whose name is already taken (new phone, cleared browser) asks his leaders to reset him.
function requestReset(code, name) {
  const w = wardByJoin_(code);
  return requestInvite(w.id, name);
}
// Leaders sign in by email: we send a sign-in link for every ward that email leads. Same answer either way.
function leaderEmailLink(email) {
  email = String(email || '').trim().toLowerCase();
  const msg = "If that email belongs to a leader, we just sent a sign-in link. Check your inbox (and spam).";
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Enter your email.');
  const c = cache_(), key = 'mail_' + email;
  if (c.get(key)) return { message: msg };
  c.put(key, '1', 300);
  const m = leaders_();
  Object.keys(m).forEach(function (t) {
    if (String(m[t].email).toLowerCase() !== email) return;
    const w = wardById_(m[t].ward);
    if (w) { try { sendLeaderLink_(m[t].email, w.name, t, 'Here is your sign-in link for FSY Daily Prep.'); } catch (x) { console.error(x); } }
  });
  return { message: msg };
}

function getData(youthCode, invite, legacy) {
  let w, me = '';
  if (invite) { const v = byInvite_(invite); w = v.ward; me = v.name; }
  else w = wardByYouth_(youthCode, legacy);
  if (!w) throw new Error('This app has no wards set up yet.');
  return { today: today_(w), me: me, youthCode: w.youth,
           ward: { name: w.name, isDefault: w.id === DEFAULT_WARD_ID },
           roster: roster_(w.id), readings: readings_(), checkins: checkins_(w.id) };
}

function checkIn(youthCode, name, thought, invite, legacy) {
  name = String(name || '').trim();
  thought = String(thought || '').trim().slice(0, 400);
  if (thought.length < 15) throw new Error('Write at least a sentence about what you read before checking in.');
  if (/^[=+\-@]/.test(thought)) thought = "'" + thought;
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    let w;
    if (invite) { const v = byInvite_(invite); w = v.ward; name = v.name; }
    else w = wardByYouth_(youthCode, legacy);
    if (roster_(w.id).indexOf(name) < 0) throw new Error('That name is not on the roster. Ask your leader to add it.');
    const t = today_(w);
    bust_(['ck_' + w.id]);
    const already = checkins_(w.id).some(function (c) { return c.date === t && c.name === name; });
    if (!already) {
      sh_('Check-ins').appendRow([new Date(), "'" + t, name, thought, w.id]);
      SpreadsheetApp.flush();
      bust_(['ck_' + w.id]);
    }
  } finally {
    lock.releaseLock();
  }
  return getData(youthCode, invite, legacy);
}

/* ---------- ward setup ---------- */
function setupInfo() { return { needsCode: !!setupCode_(), zones: ZONES }; }

function createWard(wardName, namesText, tz, setupCode, leaderName, leaderEmail) {
  tooMany_('fail_create');
  if (setupCode_()) {
    tooMany_('fail_setup');
    if (norm_(setupCode) !== norm_(setupCode_())) { failed_('fail_setup'); throw new Error("That setup code didn't work. Ask the person who shared FSY Daily Prep with you."); }
  }
  const pin = String(1000 + Math.floor(Math.random() * 9000));   // backup only; leaders sign in by email
  leaderName = clean_(leaderName, 40);
  leaderEmail = String(leaderEmail || '').trim().slice(0, 120);
  if (!leaderName) throw new Error('Enter your name.');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(leaderEmail)) throw new Error("Enter your email. It's how you get back in if you forget your PIN.");
  wardName = clean_(wardName, 60);
  if (wardName.replace(/^'/, '').length < 3) throw new Error('Enter your ward name.');
  tz = ZONES.indexOf(tz) >= 0 ? tz : TZ;
  const names = [];
  String(namesText || '').split(/[\n,]/).forEach(function (n) {
    n = clean_(n, 40);
    if (n && names.indexOf(n) < 0) names.push(n);
  });
  if (names.length > MAX_NAMES) throw new Error('Add up to ' + MAX_NAMES + ' names for now.');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    bust_(['wards']);
    if (wards_().length >= MAX_WARDS) throw new Error('This app is full right now. Contact the person who shared it.');
    const w = { id: 'w_' + Date.now().toString(36), name: wardName, youth: youthCode_(wardName), leader: pin };
    sh_('Wards').appendRow([w.id, w.name, w.youth, "'" + w.leader, "'" + Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'), tz]);
    const sh = sh_('Roster');
    if (names.length) sh.getRange(sh.getLastRow() + 1, 1, names.length, 4)
      .setValues(names.map(function (n) { return [n, '', w.id, newToken_()]; }));
    let lt;
    do { lt = newCode_(12); } while (leaders_()[lt]);
    leadersSheet_().appendRow([leaderName, w.id, lt, "'" + Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'), leaderEmail]);
    bust_(['wards', 'roster_' + w.id, 'invites', 'leaders']);
    failed_('fail_create');   // counts toward a limit of 5 new wards per 15 minutes
    try { sendLeaderLink_(leaderEmail, wardName, lt, 'Your ward is set up on FSY Daily Prep. Use this link to open your dashboard on any phone or computer.'); } catch (x) { console.error(x); }
    return { wardName: w.name.replace(/^'/, ''), cred: lt };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- leaders ---------- */
function leaderData(leaderCode) {
  const w = wardByLeader_(leaderCode);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { ensureTokens_(w.id); } finally { lock.releaseLock(); }
  const titles = {};
  readings_().forEach(function (r) { titles[r.date] = r.title; });
  const boys = rosterRows_(w.id).map(function (r) { return { name: r.name, joined: !!r.joined }; });
  return { today: today_(w), boys: boys, joinLink: joinLink_(w), tz: tzOf_(w),
           ward: { name: w.name, youthCode: w.youthPretty || w.youth, link: link_(w), isDefault: w.id === DEFAULT_WARD_ID },
           leaders: leaderList_(w.id), roster: roster_(w.id), invites: inviteList_(w.id), checkins: checkinsFull_(w.id), titles: titles };
}

function leaderAddName(leaderCode, name) {
  const w = wardByLeader_(leaderCode);
  name = clean_(name, 40);
  if (!name) throw new Error('Enter a name.');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    bust_(['roster_' + w.id]);
    const r = roster_(w.id);
    if (r.indexOf(name) >= 0) throw new Error(name + ' is already on the roster.');
    if (r.length >= MAX_NAMES) throw new Error('Your roster is full.');
    sh_('Roster').appendRow([name, '', w.id, newToken_()]);
    bust_(['roster_' + w.id, 'invites']);
  } finally {
    lock.releaseLock();
  }
  return leaderData(leaderCode);
}

function leaderAddNames(leaderCode, text) {
  const w = wardByLeader_(leaderCode);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    bust_(['roster_' + w.id]);
    const have = roster_(w.id), add = [];
    String(text || '').split(/[\n,]/).forEach(function (n) {
      n = clean_(n, 40);
      if (n && have.indexOf(n) < 0 && add.indexOf(n) < 0) add.push(n);
    });
    if (!add.length) throw new Error('Type at least one new name.');
    if (have.length + add.length > MAX_NAMES) throw new Error('That would put more than ' + MAX_NAMES + ' names on your list.');
    const sh = sh_('Roster');
    sh.getRange(sh.getLastRow() + 1, 1, add.length, 4).setValues(add.map(function (n) { return [n, '', w.id, newToken_()]; }));
    bust_(['roster_' + w.id, 'invites']);
  } finally {
    lock.releaseLock();
  }
  return leaderData(leaderCode);
}
// New phone or logged out: gives him a fresh sign-in and lets him tap his name on the join page again.
function leaderResetBoy(leaderCode, name) {
  const w = wardByLeader_(leaderCode);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    rosterRows_(w.id).forEach(function (r) {
      if (r.name !== String(name).trim()) return;
      sh_('Roster').getRange(r.row, 4, 1, 2).setValues([[newToken_(), '']]);
    });
    bust_(['invites']);
  } finally {
    lock.releaseLock();
  }
  return leaderData(leaderCode);
}

function leaderRemoveName(leaderCode, name) {
  const w = wardByLeader_(leaderCode);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = sh_('Roster');
    const vals = sh.getDataRange().getValues();
    for (let i = vals.length - 1; i >= 1; i--) {
      if (String(vals[i][0]).trim() === String(name).trim() && wardOf_(vals[i][2]) === w.id) sh.deleteRow(i + 1);
    }
    bust_(['roster_' + w.id, 'invites']);
  } finally {
    lock.releaseLock();
  }
  return leaderData(leaderCode);
}

function leaderList_(wardId) {
  const m = leaders_(), out = [];
  Object.keys(m).forEach(function (t) { if (m[t].ward === wardId) out.push({ name: m[t].name, email: m[t].email || '', link: base_() + '?L=' + t }); });
  return out.sort(function (a, b) { return a.name.localeCompare(b.name); });
}
function leaderAddLeader(leaderCode, name, email) {
  const w = wardByLeader_(leaderCode);
  name = clean_(name, 40);
  email = String(email || '').trim().slice(0, 120);
  if (!name) throw new Error("Enter the leader's name.");
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error("Enter the leader's email. That's how he signs in.");
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (leaderList_(w.id).some(function (x) { return x.name === name; })) throw new Error(name + ' already has a link.');
    let t;
    do { t = newCode_(12); } while (leaders_()[t]);
    leadersSheet_().appendRow([name, w.id, t, "'" + today_(w), email]);
    bust_(['leaders']);
    try { sendLeaderLink_(email, w.name, t, 'You were added as a leader for ' + String(w.name).replace(/[<>&]/g, '') + ' on FSY Daily Prep. Open your dashboard to see your young men\'s check-ins and responses.'); } catch (x) { console.error(x); }
  } finally {
    lock.releaseLock();
  }
  return leaderData(leaderCode);
}
function leaderRemoveLeader(leaderCode, name) {
  const w = wardByLeader_(leaderCode);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = leadersSheet_(), vals = sh.getDataRange().getValues();
    for (let i = vals.length - 1; i >= 1; i--) {
      if (String(vals[i][0]).trim() === String(name).trim() && String(vals[i][1]).trim() === w.id) sh.deleteRow(i + 1);
    }
    bust_(['leaders']);
  } finally {
    lock.releaseLock();
  }
  return leaderData(leaderCode);
}

// makes a new invite link for one young man; his old link stops working
function leaderResetInvite(leaderCode, name) {
  const w = wardByLeader_(leaderCode);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = sh_('Roster'), vals = sh.getDataRange().getValues();
    for (let i = 1; i < vals.length; i++) {
      if (String(vals[i][0]).trim() === String(name).trim() && wardOf_(vals[i][2]) === w.id) {
        sh.getRange(i + 1, 4).setValue(newToken_());
      }
    }
    bust_(['invites']);
  } finally {
    lock.releaseLock();
  }
  return leaderData(leaderCode);
}

function leaderNewCode(leaderCode, which) {
  const w = wardByLeader_(leaderCode);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const code = youthCode_(w.name);
    sh_('Wards').getRange(w.row, 3).setValue(code);
    bust_(['wards']);
    return code;
  } finally {
    lock.releaseLock();
  }
}

/* ---------- young man lost his link ---------- */
// Ward names only (no codes, no young men's names), for the "Lost your link?" picker.
function wardList() {
  return wards_().filter(function (w) { return w.id && !/^test/i.test(w.name); })
    .map(function (w) { return { id: w.id, name: String(w.name).replace(/^'/, '') }; })
    .sort(function (a, b) { return a.name.localeCompare(b.name); });
}
// Emails the ward's leaders that a young man needs his invite link again.
function requestInvite(wardId, name) {
  name = String(name || '').replace(/\s+/g, ' ').trim().slice(0, 40);
  if (name.length < 2) throw new Error('Enter your name.');
  const w = wardById_(String(wardId || ''));
  if (!w) throw new Error('Pick your ward.');
  const c = cache_(), wk = 'req_' + w.id, nk = 'req_' + w.id + '_' + name.toLowerCase();
  if (c.get(nk)) return { message: 'Your leaders already got your request. Give them a little time to text you.' };
  if (Number(c.get(wk) || 0) >= 5) throw new Error('Your leaders have several requests already. Please try again in an hour.');
  const m = leaders_(), to = [];
  Object.keys(m).forEach(function (t) { if (m[t].ward === w.id && m[t].email) to.push(m[t].email); });
  if (!to.length) return { message: "Your ward's leaders don't have an email on file. Ask one of them at church for your link." };
  const esc = function (x) { return String(x).replace(/[&<>"]/g, function (ch) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]; }); };
  MailApp.sendEmail({ to: to.join(','), name: 'FSY Daily Prep',
    subject: name + ' needs the FSY Daily Prep link',
    htmlBody: '<div style="font-family:Arial,sans-serif;font-size:15px;color:#18212b">' +
      '<p><b>' + esc(name) + '</b> asked for your quorum\'s FSY Daily Prep link.</p>' +
      '<p>Text him this link (in your quorum group chat is best):<br><a href="' + joinLink_(w) + '">' + joinLink_(w) + '</a></p>' +
      '<p>If he says his name shows "already set up," open your dashboard and tap <b>Reset</b> next to his name first.</p>' +
      '<p style="color:#5b6774;font-size:13px">Don\'t recognize the name? Ignore this. Anyone can send a request, but only you can send the link.</p></div>' });
  c.put(nk, '1', 1800);
  c.put(wk, String(Number(c.get(wk) || 0) + 1), 3600);
  return { message: 'Sent. Your leaders got an email and will text you the link. Open it and tap your name.' };
}

// "Forgot your PIN?": emails each leader in that ward (who has an email on file) his personal sign-in link.
// Always answers the same way, so it never reveals whether a ward code exists. Once per ward every 10 minutes.
function forgotPin(wardCode) {
  const code = norm_(wardCode);
  const msg = "If that ward has leaders with an email on file, we just emailed each of them his personal sign-in link. Check your inbox (and spam).";
  if (!code) throw new Error('Enter your ward code.');
  const c = cache_(), key = 'forgot_' + code;
  if (c.get(key)) return { message: msg };
  c.put(key, '1', 600);
  const w = wards_().filter(function (x) { return x.youth === code; })[0];
  if (!w) return { message: msg };
  const m = leaders_();
  Object.keys(m).forEach(function (t) {
    const l = m[t];
    if (l.ward !== w.id || !l.email) return;
    const link = base_() + '?L=' + t;
    try {
      MailApp.sendEmail({ to: l.email, name: 'FSY Daily Prep', subject: 'Your FSY Daily Prep leader link',
        htmlBody: '<div style="font-family:Arial,sans-serif;font-size:15px;color:#18212b">' +
          '<p>Someone asked to sign in to ' + String(w.name).replace(/[<>&]/g, '') + ' on FSY Daily Prep.</p>' +
          '<p><a href="' + link + '">Open your leader view</a></p>' +
          '<p>Once you\'re in, you can set a new PIN under Codes. If you didn\'t ask for this, you can ignore it.</p></div>' });
    } catch (x) { console.error(x); }
  });
  return { message: msg };
}

// Sets the ward's 4-digit leader PIN. Returns the new sign-in so this phone stays signed in.
function leaderSetPin(leaderCode, newPin) {
  const w = wardByLeader_(leaderCode);
  const pin = pin_(newPin);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    sh_('Wards').getRange(w.row, 4).setValue("'" + pin);
    bust_(['wards']);
  } finally {
    lock.releaseLock();
  }
  return { cred: norm_(w.youth) + ':' + pin, wardCode: w.youthPretty || w.youth, pin: pin };
}

/* ---------- feedback ---------- */
// The small Feedback link at the bottom of every screen lands here: saved to a Feedback tab and emailed to the app owner.
function sendFeedback(kind, message, contact, context) {
  message = String(message || '').trim().slice(0, 2000);
  if (message.length < 5) throw new Error('Tell us a little more first.');
  kind = kind === 'idea' ? 'Idea' : 'Problem';
  contact = String(contact || '').trim().slice(0, 120);
  context = String(context || '').slice(0, 300);
  const c = cache_(), n = Number(c.get('fb_count') || 0);
  if (n >= 20) throw new Error('Lots of feedback right now. Please try again in an hour.');
  c.put('fb_count', String(n + 1), 3600);
  let sh = sh_('Feedback');
  if (!sh) { sh = ss_().insertSheet('Feedback'); sh.appendRow(['When', 'Type', 'Message', 'Reply to', 'Where']); sh.setFrozenRows(1); }
  const safe = function (x) { return /^[=+\-@]/.test(x) ? "'" + x : x; };
  sh.appendRow([new Date(), kind, safe(message), safe(contact), safe(context)]);
  try {
    const owner = settings_()['Feedback email'] || '';   // set in the Settings tab
    const esc = function (x) { return String(x).replace(/[&<>"]/g, function (ch) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]; }); };
    if (owner) MailApp.sendEmail({ to: owner, name: 'FSY Daily Prep', subject: 'FSY Daily Prep ' + kind.toLowerCase() + ': ' + message.slice(0, 60),
      replyTo: /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(contact) ? contact : undefined,
      htmlBody: '<div style="font-family:Arial,sans-serif;font-size:15px;color:#18212b"><p><b>' + kind + '</b></p><p>' + esc(message).replace(/\n/g, '<br>') + '</p>' +
        '<p style="color:#5b6774;font-size:13px">Reply to: ' + (esc(contact) || 'not given') + '<br>Where: ' + esc(context) + '</p></div>' });
  } catch (x) { console.error(x); }
  return { message: 'Thanks. It went straight to the person who runs FSY Daily Prep.' };
}

// Run this from the editor after you edit the sheet by hand, so the app sees your change right away.
function clearCache() {
  const keys = ['wards', 'readings', 'invites', 'leaders', 'settings'];
  wards_().forEach(function (w) { keys.push('roster_' + w.id, 'ck_' + w.id); });
  bust_(keys);
}


/* ---------- reminders ---------- */

// fsydaily.com calls this every hour to ask which young men still need a reminder.
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (x) {}
  if (!body || !pushSecret_() || body.secret !== pushSecret_()) return json_({ error: 'not allowed' });
  if (body.action !== 'status') return json_({ error: 'unknown action' });
  const out = {};
  (body.tokens || []).slice(0, 2000).forEach(function (tok) {
    try {
      const v = byInvite_(tok), w = v.ward, t = today_(w);
      const done = checkins_(w.id).some(function (c) { return c.date === t && c.name === v.name; });
      out[tok] = { ok: true, done: done,
                   hour: Number(Utilities.formatDate(new Date(), tzOf_(w), 'H')),
                   dow: Number(Utilities.formatDate(new Date(), tzOf_(w), 'u')) % 7 };   // 0 = Sunday
    } catch (x) { out[tok] = { ok: false }; }
  });
  return json_({ status: out });
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// Runs every hour. At 6 PM in each ward's time zone, emails that ward's leaders who still hasn't checked in.
function sendLeaderSummaries() {
  const props = PropertiesService.getScriptProperties();
  const leaders = leaders_();
  wards_().forEach(function (w) {
    try {
      if (Number(Utilities.formatDate(new Date(), tzOf_(w), 'H')) !== SUMMARY_HOUR) return;
      const t = today_(w), key = 'sum_' + w.id + '_' + t;
      if (props.getProperty(key)) return;
      const to = Object.keys(leaders).filter(function (k) { return leaders[k].ward === w.id && leaders[k].email; })
        .map(function (k) { return { email: leaders[k].email, link: base_() + '?L=' + k }; });
      if (!to.length) return;
      const roster = roster_(w.id);
      if (!roster.length) return;
      bust_(['ck_' + w.id]);
      const done = {};
      checkins_(w.id).forEach(function (c) { if (c.date === t) done[c.name] = 1; });
      const missing = roster.filter(function (n) { return !done[n]; });
      const wardName = w.name.replace(/^'/, '');
      const subject = missing.length
        ? wardName + ': ' + missing.length + ' of ' + roster.length + ' still need to check in today'
        : wardName + ': everyone checked in today';
      const esc = function (x) { return String(x).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
      const html = '<div style="font-family:Arial,sans-serif;font-size:15px;color:#18212b">' +
        '<p><b>' + (roster.length - missing.length) + ' of ' + roster.length + '</b> have checked in today.</p>' +
        (missing.length ? '<p>Still to go:</p><ul>' + missing.map(function (n) { return '<li>' + esc(n) + '</li>'; }).join('') + '</ul>' : '<p>Nice work, everyone is in.</p>') +
        '<p><a href="{LINK}" style="display:inline-block;background:#1d3047;color:#fff;text-decoration:none;padding:10px 16px;border-radius:8px;font-weight:bold">Open my dashboard</a></p>' +
        '<p style="color:#5b6774;font-size:13px">That button is your personal sign-in link, so don\'t forward this email. You get this because a leader in your ward added your email in FSY Daily Prep.</p></div>';
      to.forEach(function (t) {
        MailApp.sendEmail({ to: t.email, subject: subject, htmlBody: html.replace('{LINK}', t.link), name: 'FSY Daily Prep' });
      });
      props.setProperty(key, '1');
    } catch (x) { console.error(w.id, x); }
  });
}

// Run this once from the editor (select it, click Run) to start the hourly leader emails and automatic updates.
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (tr) {
    const f = tr.getHandlerFunction();
    if (f === 'sendLeaderSummaries' || f === 'autoUpdate') ScriptApp.deleteTrigger(tr);
  });
  ScriptApp.newTrigger('sendLeaderSummaries').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('autoUpdate').timeBased().everyHours(1).create();
}

/* ---------- automatic updates ---------- */
// The latest code lives in GitHub. Every hour this checks for a new version and, if there is one,
// installs it and republishes the web app at the same link. Run updateApp from the editor to do it right now.
const REPO_RAW = 'https://raw.githubusercontent.com/Crosstownpest/fsydaily/main/apps-script/';
const DEPLOYMENT_ID = 'AKfycbywBQ2vv05hkwKkOFd384q933YG2rEtsgNnisQ6FC5TYrqA_NURZr4hhgdUZxyIqNmkDg';

function updateApp() { const r = selfUpdate_(true); console.log(r); return r; }
function autoUpdate() { try { selfUpdate_(false); } catch (e) { console.error('Auto update failed: ' + e.message); } }

function selfUpdate_(force) {
  const files = [['Code', 'SERVER_JS', 'Code.gs'], ['Index', 'HTML', 'Index.html'], ['appsscript', 'JSON', 'appsscript.json']];
  const got = files.map(function (f) {
    const r = UrlFetchApp.fetch(REPO_RAW + f[2] + '?t=' + Date.now(), { muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) throw new Error("Couldn't download " + f[2] + ' (' + r.getResponseCode() + ')');
    return { name: f[0], type: f[1], source: r.getContentText() };
  });
  if (got[0].source.indexOf('FSY_DAILY_PREP_CODE') < 0 || got[1].source.indexOf('</html>') < 0) throw new Error('The downloaded files look incomplete. Nothing was changed.');
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, got.map(function (g) { return g.source; }).join('\u0000'));
  const hash = Utilities.base64Encode(digest);
  const props = PropertiesService.getScriptProperties();
  if (!force && props.getProperty('codeHash') === hash) return 'Already up to date.';

  const id = ScriptApp.getScriptId(), api = 'https://script.googleapis.com/v1/projects/' + id;
  function call(method, url, body) {
    const r = UrlFetchApp.fetch(url, { method: method, contentType: 'application/json', payload: JSON.stringify(body),
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    if (r.getResponseCode() >= 300) throw new Error(method.toUpperCase() + ' ' + url.split('/').pop() + ' failed: ' + r.getContentText().slice(0, 300));
    return JSON.parse(r.getContentText() || '{}');
  }
  call('put', api + '/content', { files: got });
  const v = call('post', api + '/versions', { description: 'Update ' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm') });
  call('put', api + '/deployments/' + DEPLOYMENT_ID, { deploymentConfig: {
    scriptId: id, versionNumber: v.versionNumber, manifestFileName: 'appsscript', description: 'FSY Daily Prep' } });
  props.setProperty('codeHash', hash);
  return 'Updated to version ' + v.versionNumber + '.';
}
