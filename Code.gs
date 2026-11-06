/**
 * PQL office hours – booking web app
 * Bound to the "OH application" spreadsheet (Settings, PQLs, Bookings tabs).
 */

const TAB_SETTINGS = 'Settings';
const TAB_PQLS = 'PQLs';
const TAB_BOOKINGS = 'Bookings';
const TAB_HOURS = 'Hours';
const CACHE_KEY = 'state_v1';
const CACHE_SECONDS = 60;

// Labels in column A of the Settings tab.
const KEY = {
  title: 'title',
  start: 'start',
  end: 'end',
  len: 'slot length (hours)',
  cap: 'pqls per slot',
  max: 'max slots per pql',
  link: 'meeting link',
  password: 'admin password'
};

const HOUR = 3600 * 1000;

function doGet(e) {
  if (e && e.parameter && e.parameter.fn) {
    let args = [];
    try { args = JSON.parse(e.parameter.args || '[]'); } catch (x) {}
    return api_({ fn: e.parameter.fn, args: args });
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('PQL office hours')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ------------------------------------------------------------------ */
/* JSON API (used by the page, works without a Google login)          */
/* ------------------------------------------------------------------ */

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (x) { return json_({ ok: false, error: 'Bad request' }); }
  return api_(req);
}

function api_(req) {
  const fns = { getState: getState, book: book, unbook: unbook, adminLogin: adminLogin, adminSave: adminSave };
  try {
    const f = fns[req && req.fn];
    if (!f) throw new Error('Unknown action');
    return json_({ ok: true, data: f.apply(null, req.args || []) });
  } catch (err) {
    return json_({ ok: false, error: err.message });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------------------------------------------ */
/* Reading the sheet                                                  */
/* ------------------------------------------------------------------ */

function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('Missing tab: ' + name);
  return sh;
}

/** Returns { key: { row, value } } for every label in Settings column A. */
function settingsMap_() {
  const values = sheet_(TAB_SETTINGS).getDataRange().getValues();
  const map = {};
  for (let r = 1; r < values.length; r++) {
    const label = String(values[r][0] || '').trim().toLowerCase();
    if (label) map[label] = { row: r + 1, value: values[r][1] };
  }
  return map;
}

function toDate_(v) {
  if (v instanceof Date) return v;
  const s = String(v || '').trim();
  if (!s) return null;
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  return Utilities.parseDate(s.replace('T', ' ').slice(0, 16), tz, 'yyyy-MM-dd HH:mm');
}

function settings_() {
  const m = settingsMap_();
  const get = k => (m[k] ? m[k].value : '');
  const start = toDate_(get(KEY.start));
  const end = toDate_(get(KEY.end));
  if (!start || !end) throw new Error('Set Start and End on the Settings tab.');
  return {
    title: String(get(KEY.title) || 'PQL office hours'),
    start: start.getTime(),
    end: end.getTime(),
    len: Number(get(KEY.len)) || 2,
    cap: Number(get(KEY.cap)) || 0,
    max: Number(get(KEY.max)) || 0,
    link: String(get(KEY.link) || ''),
    password: String(get(KEY.password) || '')
  };
}

function displayName_(email) {
  return email.split('@')[0].split(/[._-]+/)
    .map(w => w.replace(/\d+/g, ''))
    .filter(Boolean)
    .map(w => w.charAt(0).toUpperCase() + w.slice(1))
    .join(' ');
}

function pqls_() {
  const values = sheet_(TAB_PQLS).getDataRange().getValues();
  const seen = {};
  const out = [];
  for (let r = 1; r < values.length; r++) {
    const email = String(values[r][0] || '').trim().toLowerCase();
    if (!email || seen[email]) continue;
    seen[email] = true;
    out.push({ email: email, name: String(values[r][1] || '').trim() || displayName_(email) });
  }
  return out;
}

/** Each booking: { row, start (ms), email } */
function bookings_() {
  const values = sheet_(TAB_BOOKINGS).getDataRange().getValues();
  const out = [];
  for (let r = 1; r < values.length; r++) {
    const start = values[r][0];
    const email = String(values[r][2] || '').trim().toLowerCase();
    if (!(start instanceof Date) || !email) continue;
    const at = values[r][4] instanceof Date ? values[r][4].getTime() : start.getTime();
    out.push({ row: r + 1, start: start.getTime(), email: email, bookedAt: at });
  }
  return out;
}

function slotList_(s) {
  const out = [];
  for (let t = s.start; t + s.len * HOUR <= s.end; t += s.len * HOUR) {
    out.push({ start: t, end: t + s.len * HOUR });
  }
  return out;
}

const minuteKey_ = ms => Math.round(ms / 60000);

/* ------------------------------------------------------------------ */
/* Called from the page                                               */
/* ------------------------------------------------------------------ */

/** Reads the sheet once and returns { st, rows }. */
function read_() {
  const s = settings_();
  const pqls = pqls_();
  const valid = {};
  pqls.forEach(p => (valid[p.email] = true));
  const slots = slotList_(s).map(x => ({ start: x.start, end: x.end, who: [] }));
  const byKey = {};
  slots.forEach(x => (byKey[minuteKey_(x.start)] = x));
  const rows = bookings_();
  rows.forEach(b => {
    const slot = byKey[minuteKey_(b.start)];
    if (slot && valid[b.email] && slot.who.indexOf(b.email) < 0) slot.who.push(b.email);
  });
  const st = {
    title: s.title, start: s.start, end: s.end, len: s.len, cap: s.cap, max: s.max, link: s.link,
    pqls: pqls, slots: slots
  };
  return { st: st, rows: rows };
}

function cachePut_(st) {
  try { CacheService.getScriptCache().put(CACHE_KEY, JSON.stringify(st), CACHE_SECONDS); } catch (e) {}
}

function cacheClear_() {
  try { CacheService.getScriptCache().remove(CACHE_KEY); } catch (e) {}
}

function stamp_(st) {
  st.serverNow = Date.now();
  return st;
}

/** Page load and polling: served from a short cache when possible. */
function getState() {
  let st = null;
  try {
    const hit = CacheService.getScriptCache().get(CACHE_KEY);
    if (hit) st = JSON.parse(hit);
  } catch (e) {}
  if (!st) {
    st = read_().st;
    cachePut_(st);
  }
  return stamp_(st);
}

function book(slotStart, email) {
  email = String(email || '').trim().toLowerCase();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const st = read_().st;
    const pql = st.pqls.filter(p => p.email === email)[0];
    if (!pql) throw new Error("That email isn't on the PQL list.");
    const slot = st.slots.filter(x => minuteKey_(x.start) === minuteKey_(slotStart))[0];
    if (!slot) throw new Error('That slot no longer exists. Refresh the page.');
    if (slot.end <= Date.now()) throw new Error('That slot has already ended.');
    if (slot.who.indexOf(email) >= 0) throw new Error("You're already in this slot.");
    if (st.cap && slot.who.length >= st.cap) throw new Error('This slot just filled up. Pick another.');
    const mine = st.slots.filter(x => x.who.indexOf(email) >= 0).length;
    if (st.max && mine >= st.max) {
      throw new Error(pql.name + ' already has ' + mine + ' slot' + (mine > 1 ? 's' : '') + ' (limit ' + st.max + ').');
    }
    sheet_(TAB_BOOKINGS).appendRow([new Date(slot.start), new Date(slot.end), email, pql.name, new Date()]);
    slot.who.push(email);
    cachePut_(st);
    return stamp_(st);
  } finally {
    lock.releaseLock();
  }
}

function unbook(slotStart, email) {
  email = String(email || '').trim().toLowerCase();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const data = read_();
    const st = data.st;
    const slot = st.slots.filter(x => minuteKey_(x.start) === minuteKey_(slotStart))[0];
    if (slot && slot.end <= Date.now()) throw new Error("That slot has ended, so it can't be changed.");
    const rows = data.rows
      .filter(b => b.email === email && minuteKey_(b.start) === minuteKey_(slotStart))
      .map(b => b.row)
      .sort((a, b) => b - a);
    const sh = sheet_(TAB_BOOKINGS);
    rows.forEach(r => sh.deleteRow(r));
    if (slot) slot.who = slot.who.filter(e => e !== email);
    cachePut_(st);
    return stamp_(st);
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------------------------------------------ */
/* Hours tab (updated hourly by a timer)                              */
/* ------------------------------------------------------------------ */

/**
 * Hours per PQL = time actually covered: from when they joined (Booked at,
 * or the slot start if they joined before it) until the slot end, counting
 * only time that has already passed.
 */
function updateHours() {
  const data = read_();
  const now = Date.now();
  const slotByKey = {};
  data.st.slots.forEach(x => (slotByKey[minuteKey_(x.start)] = x));
  const totals = {};
  const counted = {};
  data.rows.forEach(b => {
    const slot = slotByKey[minuteKey_(b.start)];
    if (!slot) return;
    const key = b.email + '|' + minuteKey_(b.start);
    if (counted[key]) return;
    counted[key] = true;
    const from = Math.max(slot.start, b.bookedAt);
    const to = Math.min(slot.end, now);
    if (to > from) totals[b.email] = (totals[b.email] || 0) + (to - from);
  });
  const rows = data.st.pqls.map(p => [p.email, p.name, Math.round((totals[p.email] || 0) / HOUR * 100) / 100]);
  let sh = SpreadsheetApp.getActive().getSheetByName(TAB_HOURS);
  if (!sh) sh = SpreadsheetApp.getActive().insertSheet(TAB_HOURS);
  sh.getRange(1, 1, 1, 4).setValues([['Email', 'Name', 'Hours spent', 'Last updated']]);
  const last = Math.max(sh.getLastRow(), 2);
  sh.getRange(2, 1, last - 1, 4).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 3).setValues(rows);
  sh.getRange(2, 4).setValue(new Date(now));
}

/** Run once from the editor: makes updateHours run every hour. */
function installHourlyTimer() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'updateHours')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('updateHours').timeBased().everyHours(1).create();
  updateHours();
}

/* ------------------------------------------------------------------ */
/* Admin                                                              */
/* ------------------------------------------------------------------ */

function checkPw_(pw) {
  if (String(pw || '') !== settings_().password) throw new Error("That password isn't right.");
}

function adminLogin(pw) {
  checkPw_(pw);
  const st = read_().st;
  cachePut_(st);
  return stamp_(st);
}

/**
 * cfg: { title, start (ms), end (ms), len, cap, max, link, emails: [..] }
 */
function adminSave(pw, cfg) {
  checkPw_(pw);
  const start = Number(cfg.start), end = Number(cfg.end), len = Number(cfg.len);
  if (!(end > start)) throw new Error('End time must be after start time.');
  if ((end - start) / HOUR < len) throw new Error('The window is shorter than one slot.');
  if (!/^https?:\/\//.test(String(cfg.link || ''))) throw new Error('Add a meeting link starting with https://');
  const emails = (cfg.emails || []).map(e => String(e).trim().toLowerCase())
    .filter((e, i, a) => /^[^@\s]+@[^@\s]+\.[a-z]{2,}$/.test(e) && a.indexOf(e) === i);
  if (!emails.length) throw new Error('Paste at least one PQL email.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const m = settingsMap_();
    const sh = sheet_(TAB_SETTINGS);
    const put = (k, v) => { if (m[k]) sh.getRange(m[k].row, 2).setValue(v); };
    put(KEY.title, String(cfg.title || 'PQL office hours'));
    put(KEY.start, new Date(start));
    put(KEY.end, new Date(end));
    put(KEY.len, len);
    put(KEY.cap, Number(cfg.cap) || 0);
    put(KEY.max, Number(cfg.max) || 0);
    put(KEY.link, String(cfg.link));

    // Rewrite only column A of PQLs, so the display-name formula in column B keeps working.
    const p = sheet_(TAB_PQLS);
    const last = Math.max(p.getLastRow(), 2);
    p.getRange(2, 1, last - 1, 1).clearContent();
    p.getRange(2, 1, emails.length, 1).setValues(emails.map(e => [e]));
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  cacheClear_();
  const st = read_().st;
  cachePut_(st);
  return stamp_(st);
}
