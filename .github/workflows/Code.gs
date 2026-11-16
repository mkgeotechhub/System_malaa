/**
 * MALAA POULTRY SYSTEM - Google Sheets receiver (v2)
 *
 * What it does
 *  - Receives records from dashboard.html (Settings > Google Sheets sync).
 *  - Creates one tab per module the first time a record arrives:
 *    Batches, Egg Production, Mortality, Orders, Sales, Expenses, Inventory.
 *  - Matches rows by the ID in column A, so edits update the same row and deletes remove it.
 *  - Writes every change to a "Sync Log" tab (time, action, tab, ID).
 *  - Only records the admin has approved are ever sent by the dashboard.
 *
 * Sign-in (used by login.html and the dashboard)
 *  - Users live in a "Users" tab of this Google Sheet. Passwords are stored as salted SHA-256
 *    hashes, never as plain text. Roles: admin, boss, assistant, staff.
 *  - login returns a session token (valid about 6 hours). Only a signed-in admin can write records.
 *  - 5 wrong passwords for one username locks it for 5 minutes.
 *
 * Setup
 *  1. Extensions > Apps Script, paste this file.
 *  2. Nothing to run: the first administrator (mokua@malaa) is created automatically on the first
 *     sign-in. (createAdmin() below is optional, to reset an admin password.)
 *  3. Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone.
 *  4. Paste the /exec URL into SCRIPT_URL at the top of login.html.
 * After editing this code later: Deploy > Manage deployments > Edit > New version.
 */
const HEAD_COLOR = '#2F7D32';
const LOG_TAB = 'Sync Log';
const TABS = ['Batches', 'Egg Production', 'Mortality', 'Orders', 'Sales', 'Expenses', 'Inventory'];
const MAX_LOG_ROWS = 5000;
const MAX_COLS = 60;
const MAX_BULK = 200;
const USERS_TAB = 'Users';
const USER_HEAD = ['Username', 'Name', 'Role', 'Active', 'Salt', 'Hash', 'Created'];
const ROLE_LIST = ['admin', 'boss', 'assistant', 'staff'];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('MALAA')
    .addItem('Create all tabs', 'createAllTabs')
    .addToUi();
}

/** Creates any missing module tabs (headers are written when the first record arrives). */
function createAllTabs() {
  const ss = SpreadsheetApp.getActive();
  TABS.concat([LOG_TAB]).forEach(function (n) { if (!ss.getSheetByName(n)) ss.insertSheet(n); });
  SpreadsheetApp.getActive().toast('Tabs are ready.');
}

function doGet() {
  return out({ ok: true, message: 'MALAA receiver is running. Send data from the dashboard.' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    if (!e || !e.postData || !e.postData.contents) return out({ ok: false, error: 'No data received' });
    const b = JSON.parse(e.postData.contents);
    if (b.action === 'login') return out(login_(b));
    if (b.action === 'verify') {
      const s = session_(b.token);
      return out(s ? { ok: true, username: s.username, name: s.name, role: s.role } : { ok: false, error: 'Session expired. Please sign in again.' });
    }
    if (b.action === 'savepdf') return out(savePdf_(b));
    if (b.action === 'ask_ai') return out(askAi_(b));
    if (b.action === 'heartbeat') return out(heartbeat_(b));
    if (b.action === 'logout') {
      if (b.token) {
        const c = CacheService.getScriptCache(), u = c.get('t_' + b.token);
        c.remove('t_' + b.token);
        if (u) withLock_(function () { const p = readPresence_(); delete p[u]; c.put(PRESENCE_KEY, JSON.stringify(p), 600); });
      }
      return out({ ok: true });
    }
    if (b.action === 'password') return out(withLock_(function () { return changePw_(b); }));
    if (String(b.action).indexOf('users.') === 0) return out(withLock_(function () { return usersAction_(b); }));

    if (b.action === 'ping') {
      return out({ ok: true, message: 'Connected to "' + SpreadsheetApp.getActive().getName() + '"' });
    }
    if (TABS.indexOf(b.sheet) < 0) return out({ ok: false, error: 'Unknown tab: ' + b.sheet });

    const s = session_(b.token);
    if (!s || s.role !== 'admin') return out({ ok: false, error: 'Only a signed-in administrator can write to the sheet. Please sign in again.' });
    lock.waitLock(20000);
    if (b.action === 'upsert') {
      const sh = getSheet(b.sheet, b.header);
      checkRow(b.row);
      upsertMany(sh, [b.row]);
      log('upsert', b.sheet, b.row[0]);
    } else if (b.action === 'bulk') {
      if (!b.rows || b.rows.length > MAX_BULK) return out({ ok: false, error: 'Bulk limit is ' + MAX_BULK + ' rows per request' });
      const sh = getSheet(b.sheet, b.header);
      b.rows.forEach(checkRow);
      upsertMany(sh, b.rows);
      log('bulk', b.sheet, b.rows.length + ' rows');
    } else if (b.action === 'delete') {
      const sh = getSheet(b.sheet, null);
      removeRow(sh, b.id);
      log('delete', b.sheet, b.id);
    } else {
      return out({ ok: false, error: 'Unknown action: ' + b.action });
    }
    return out({ ok: true });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function checkRow(row) {
  if (!row || !row.length || row.length > MAX_COLS) throw new Error('Invalid row');
}

/** Stops text that starts with = + - @ from being treated as a formula. */
function safe(v) {
  return (typeof v === 'string' && /^[=+\-@]/.test(v)) ? "'" + v : v;
}

function getSheet(name, header) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (header && header.length) {
    if (header.length > MAX_COLS) throw new Error('Header too long');
    const cur = sh.getRange(1, 1, 1, header.length).getValues()[0];
    if (cur.join('|') !== header.join('|')) {
      sh.getRange(1, 1, 1, header.length).setValues([header.map(safe)])
        .setFontWeight('bold').setBackground(HEAD_COLOR).setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

function upsertMany(sh, rows) {
  if (!rows || !rows.length) return;
  const last = sh.getLastRow();
  const ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]); }) : [];
  const appends = [];
  rows.forEach(function (raw) {
    const row = raw.map(safe);
    const i = ids.indexOf(String(row[0]));
    if (i >= 0) sh.getRange(i + 2, 1, 1, row.length).setValues([row]);
    else { appends.push(row); ids.push(String(row[0])); }
  });
  if (appends.length) {
    const width = Math.max.apply(null, appends.map(function (r) { return r.length; }));
    appends.forEach(function (r) { while (r.length < width) r.push(''); });
    sh.getRange(sh.getLastRow() + 1, 1, appends.length, width).setValues(appends);
  }
}

function removeRow(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]); });
  const i = ids.indexOf(String(id));
  if (i >= 0) sh.deleteRow(i + 2);
}

function log(action, tab, id) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(LOG_TAB) || ss.insertSheet(LOG_TAB);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 4).setValues([['Time', 'Action', 'Tab', 'ID']])
      .setFontWeight('bold').setBackground(HEAD_COLOR).setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  sh.appendRow([new Date(), action, tab, safe(String(id))]);
  const over = sh.getLastRow() - MAX_LOG_ROWS;
  if (over > 0) sh.deleteRows(2, over);
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- users and sign-in ---------------- */
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function usersSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(USERS_TAB);
  if (!sh) {
    sh = ss.insertSheet(USERS_TAB);
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange('E:F').setNumberFormat('@');
    sh.getRange(1, 1, 1, USER_HEAD.length).setValues([USER_HEAD])
      .setFontWeight('bold').setBackground(HEAD_COLOR).setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function readUsers_() {
  const sh = usersSheet_(), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, USER_HEAD.length).getValues().map(function (r, i) {
    return { row: i + 2, username: String(r[0]), name: String(r[1]), role: String(r[2]),
      active: r[3] === true || String(r[3]).toUpperCase() === 'TRUE', salt: String(r[4]), hash: String(r[5]), created: String(r[6]) };
  });
}

function hash_(salt, pw) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pw, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}

function writeUser_(u) {
  const sh = usersSheet_();
  const vals = [[u.username, u.name, u.role, u.active, u.salt, u.hash, u.created]];
  if (u.row) sh.getRange(u.row, 1, 1, USER_HEAD.length).setValues(vals);
  else sh.getRange(sh.getLastRow() + 1, 1, 1, USER_HEAD.length).setValues(vals);
}

/* First-run admin (mokua@malaa). Created automatically the first time anyone tries to sign in
 * while the Users tab is empty. Only a salted hash is stored here, never the password. */
const SEED_USER = { username: 'mokua@malaa', name: 'Mokua (Admin)', role: 'admin', active: true, salt: 'sSeed7f3a91c2d4e', hash: 'd64616d072974fa9eb2fb85799353a90b9f791f22f625a029342604cd0162bc4' };

function ensureSeedAdmin_() {
  if (readUsers_().length) return;
  withLock_(function () {
    if (readUsers_().length) return;
    writeUser_({ username: SEED_USER.username, name: SEED_USER.name, role: SEED_USER.role, active: true,
      salt: SEED_USER.salt, hash: SEED_USER.hash, created: Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd') });
  });
}

function login_(b) {
  ensureSeedAdmin_();
  const uname = String(b.username || '').trim().toLowerCase();
  const cache = CacheService.getScriptCache(), fk = 'f_' + uname;
  if (Number(cache.get(fk) || 0) >= 5) return { ok: false, error: 'Too many attempts. Try again in 5 minutes.' };
  const rec = readUsers_().filter(function (x) { return x.username.toLowerCase() === uname; })[0];
  const good = rec && rec.active && hash_(rec.salt, String(b.password || '')) === rec.hash;
  if (!good) {
    cache.put(fk, String(Number(cache.get(fk) || 0) + 1), 300);
    return { ok: false, error: rec && !rec.active ? 'This account is deactivated. Ask the administrator.' : 'Wrong username or password.' };
  }
  cache.remove(fk);
  const token = Utilities.getUuid() + Utilities.getUuid();
  cache.put('t_' + token, rec.username, 21600);
  return { ok: true, token: token, username: rec.username, name: rec.name, role: rec.role };
}

/** Returns the current user for a valid token, re-checked against the Users tab. */
function session_(token) {
  if (!token) return null;
  const uname = CacheService.getScriptCache().get('t_' + token);
  if (!uname) return null;
  const rec = readUsers_().filter(function (x) { return x.username === uname; })[0];
  return rec && rec.active ? { username: rec.username, name: rec.name, role: rec.role } : null;
}

function adminCount_(users) {
  return users.filter(function (x) { return x.role === 'admin' && x.active; }).length;
}

function usersAction_(b) {
  const s = session_(b.token);
  if (!s || s.role !== 'admin') return { ok: false, error: 'Administrator access required.' };
  const users = readUsers_();
  const un = String(b.username || '').trim().toLowerCase();
  const rec = users.filter(function (x) { return x.username.toLowerCase() === un; })[0];
  if (b.action === 'users.list') {
    return { ok: true, users: users.map(function (x) { return { username: x.username, name: x.name, role: x.role, active: x.active, created: x.created }; }) };
  }
  if (b.action === 'users.save') {
    const name = String(b.name || '').trim(), role = String(b.role || ''), pw = String(b.password || '');
    if (!name) return { ok: false, error: 'Full name is required.' };
    if (ROLE_LIST.indexOf(role) < 0) return { ok: false, error: 'Invalid role.' };
    if (!rec && !/^[a-z0-9._@-]{3,30}$/.test(un)) return { ok: false, error: 'Username: 3-30 letters, numbers, . _ @ -' };
    if ((!rec || pw) && pw.length < 6) return { ok: false, error: 'Password must be at least 6 characters.' };
    if (rec && rec.role === 'admin' && role !== 'admin' && rec.active && adminCount_(users) < 2) return { ok: false, error: 'At least one active administrator is required.' };
    const u = rec || { username: un, active: true, created: Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd') };
    u.name = name; u.role = role;
    if (!rec || pw) { u.salt = 's' + Utilities.getUuid().replace(/-/g, '').slice(0, 16); u.hash = hash_(u.salt, pw); }
    writeUser_(u);
    return { ok: true };
  }
  if (!rec) return { ok: false, error: 'User not found.' };
  if (rec.username === s.username) return { ok: false, error: 'You cannot change or delete your own account here.' };
  if (rec.role === 'admin' && rec.active && adminCount_(users) < 2) return { ok: false, error: 'At least one active administrator is required.' };
  if (b.action === 'users.toggle') { rec.active = !rec.active; writeUser_(rec); return { ok: true }; }
  if (b.action === 'users.delete') { usersSheet_().deleteRow(rec.row); return { ok: true }; }
  return { ok: false, error: 'Unknown action: ' + b.action };
}

function changePw_(b) {
  const s = session_(b.token);
  if (!s) return { ok: false, error: 'Session expired. Please sign in again.' };
  const rec = readUsers_().filter(function (x) { return x.username === s.username; })[0];
  if (hash_(rec.salt, String(b.oldPassword || '')) !== rec.hash) return { ok: false, error: 'Current password is wrong.' };
  const np = String(b.newPassword || '');
  if (np.length < 6) return { ok: false, error: 'New password must be at least 6 characters.' };
  rec.salt = 's' + Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  rec.hash = hash_(rec.salt, np);
  writeUser_(rec);
  return { ok: true };
}

/**
 * Run ONCE from the editor to create (or reset) the first administrator.
 * Type the password below, run, then delete it from the code again.
 */
function createAdmin() {
  const USERNAME = 'mokua@malaa', NAME = 'Mokua (Admin)', PASSWORD = 'TYPE_PASSWORD_HERE';
  if (PASSWORD === 'TYPE_PASSWORD_HERE') throw new Error('Type your password in PASSWORD, run again, then remove it from the code.');
  const rec = readUsers_().filter(function (x) { return x.username.toLowerCase() === USERNAME; })[0];
  const u = rec || { username: USERNAME, created: Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd') };
  u.name = NAME; u.role = 'admin'; u.active = true;
  u.salt = 's' + Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  u.hash = hash_(u.salt, PASSWORD);
  writeUser_(u);
}

/** Run from the editor, then open View > Logs. Shows who can sign in (never shows passwords). */
function diagnose() {
  const list = readUsers_().map(function (u) {
    return { username: u.username, role: u.role, active: u.active, saltOk: u.salt.length > 0, hashLength: u.hash.length };
  });
  Logger.log(list.length ? JSON.stringify(list, null, 2) : 'The Users tab has no users. Run createAdmin() first.');
}

/** Saves a PDF sent by the dashboard to Drive (folder "MALAA Receipts") and returns a view link.
 *  Needed because MIT App Inventor's WebViewer cannot download files. */
function savePdf_(b) {
  const s = session_(b.token);
  if (!s) return { ok: false, error: 'Session expired. Please sign in again.' };
  let name = String(b.name || 'document.pdf').replace(/[^A-Za-z0-9._ -]/g, '_').slice(0, 80);
  if (!/\.pdf$/i.test(name)) name += '.pdf';
  const data = String(b.data || '');
  if (!data) return { ok: false, error: 'No PDF data received' };
  if (data.length > 3000000) return { ok: false, error: 'The PDF is too large' };
  let f;
  try {
    const blob = Utilities.newBlob(Utilities.base64Decode(data), 'application/pdf', name);
    const it = DriveApp.getFoldersByName('MALAA Receipts');
    const folder = it.hasNext() ? it.next() : DriveApp.createFolder('MALAA Receipts');
    f = folder.createFile(blob);
  } catch (e) {
    return { ok: false, error: 'Google Drive permission is missing. In Apps Script run authorizeDrive once, then deploy a New version. (' + e.message + ')' };
  }
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { /* some Workspace domains block link sharing */ }
  return { ok: true, url: f.getUrl() };
}

/** Run this ONCE from the editor so Google asks for Drive permission (needed to save PDFs).
 *  Then Deploy > Manage deployments > Edit > New version > Deploy. */
function authorizeDrive() {
  const it = DriveApp.getFoldersByName('MALAA Receipts');
  if (!it.hasNext()) DriveApp.createFolder('MALAA Receipts');
  Logger.log('Drive permission is OK. Now deploy a New version.');
}

/* ---------------- who is online ---------------- */
const PRESENCE_KEY = 'presence';
const ONLINE_WINDOW_MS = 90000;

function readPresence_() {
  try { return JSON.parse(CacheService.getScriptCache().get(PRESENCE_KEY) || '{}'); } catch (e) { return {}; }
}

/** Each open dashboard calls this every ~30 seconds. Returns everyone seen in the last 90 seconds. */
function heartbeat_(b) {
  const s = session_(b.token);
  if (!s) return { ok: false, error: 'Session expired. Please sign in again.' };
  return withLock_(function () {
    const now = Date.now(), p = readPresence_(), online = [];
    p[s.username] = { name: s.name, role: s.role, ts: now };
    Object.keys(p).forEach(function (u) {
      if (now - p[u].ts > ONLINE_WINDOW_MS) delete p[u];
      else online.push({ username: u, name: p[u].name, role: p[u].role, secondsAgo: Math.round((now - p[u].ts) / 1000) });
    });
    CacheService.getScriptCache().put(PRESENCE_KEY, JSON.stringify(p), 600);
    online.sort(function (a, c) { return a.name < c.name ? -1 : 1; });
    return { ok: true, online: online };
  });
}

/* ---------------- AI assistant (Claude) ---------------- */
/* Setup: Project Settings (gear icon) > Script Properties > Add property
 *   ANTHROPIC_KEY = your Anthropic API key   (optional: AI_MODEL = model name)
 * Then run authorizeAi() once from the editor, and Deploy > Manage deployments > Edit > New version > Deploy. */
const AI_DEFAULT_MODEL = 'claude-sonnet-5-5';
const AI_MAX_PER_HOUR = 40;      // questions per user per hour
const AI_MAX_SYSTEM = 40000;     // characters of farm data accepted
const AI_MAX_MESSAGES = 12;

function askAi_(b) {
  const s = session_(b.token);
  if (!s) return { ok: false, error: 'Session expired. Please sign in again.' };
  const props = PropertiesService.getScriptProperties();
  const key = props.getProperty('ANTHROPIC_KEY');
  if (!key) return { ok: false, error: 'The assistant is not set up yet: add ANTHROPIC_KEY in Apps Script > Project Settings > Script Properties.' };

  const cache = CacheService.getScriptCache(), ck = 'ai_' + s.username;
  const used = Number(cache.get(ck) || 0);
  if (used >= AI_MAX_PER_HOUR) return { ok: false, error: 'Hourly question limit reached. Please try again later.' };

  const system = String(b.system || '').slice(0, AI_MAX_SYSTEM);
  const msgs = (Array.isArray(b.messages) ? b.messages : []).slice(-AI_MAX_MESSAGES).map(function (m) {
    return { role: m.role === 'assistant' ? 'assistant' : 'user', content: String(m.content || '').slice(0, 4000) };
  });
  if (!msgs.length || msgs[msgs.length - 1].role !== 'user') return { ok: false, error: 'No question received.' };

  let res;
  try {
    res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({ model: props.getProperty('AI_MODEL') || AI_DEFAULT_MODEL, max_tokens: 1000, system: system, messages: msgs }),
      muteHttpExceptions: true
    });
  } catch (e) {
    return { ok: false, error: 'Could not reach Anthropic: ' + e.message + ' (if this mentions permissions, run authorizeAi once).' };
  }
  let j = {};
  try { j = JSON.parse(res.getContentText()); } catch (e) {}
  if (res.getResponseCode() !== 200) return { ok: false, error: (j.error && j.error.message) || ('Anthropic error ' + res.getResponseCode()) };
  cache.put(ck, String(used + 1), 3600);
  const reply = (j.content || []).map(function (x) { return x.text || ''; }).join('');
  return reply ? { ok: true, reply: reply } : { ok: false, error: 'Empty reply from the assistant.' };
}

/** Run this ONCE from the editor so Google asks for permission to call Anthropic. */
function authorizeAi() {
  UrlFetchApp.fetch('https://api.anthropic.com/', { muteHttpExceptions: true });
  Logger.log('Permission granted. Now Deploy > Manage deployments > Edit > New version > Deploy.');
}
