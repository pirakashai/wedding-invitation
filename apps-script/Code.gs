/**
 * Wedding portal API — Pirakash & Shageepana
 * Paste into the "Wedding Portal DB" sheet: Extensions → Apps Script.
 * Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone.
 *
 * CHANGE THIS PASSWORD before deploying. It is the portal login for both of you.
 */
const PASSWORD = 'change-me';

const COLS = ['code','name','side','group','invitedTo','max','phone','notes','sentAt',
  'w_status','w_veg','w_nonveg','w_message','w_source','w_at',
  'r_status','r_veg','r_nonveg','r_message','r_source','r_at','updatedAt','rsvpBy'];
const REPO = 'pirakashai/wedding-invitation';

function doGet(e) { return handle_(e.parameter || {}); }
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); } catch (err) {}
  return handle_(body);
}

function handle_(req) {
  let out;
  try {
    const a = req.action;
    if (a === 'guest') out = guestInfo_(req.code);
    else if (a === 'rsvp') out = rsvp_(req);
    else {
      if (String(req.pw || '') !== PASSWORD) throw new Error('Wrong password');
      if (a === 'login') out = { ok: true };
      else if (a === 'list') out = { ok: true, guests: readGuests_(), settings: readSettings_() };
      else if (a === 'save') out = { ok: true, guest: saveGuest_(req.guest || {}) };
      else if (a === 'delete') out = { ok: true, deleted: deleteGuest_(req.code) };
      else if (a === 'markSent') out = { ok: true, guest: patchGuest_(req.code, { sentAt: new Date().toISOString() }) };
      else if (a === 'settings') out = { ok: true, settings: saveSettings_(req.settings || {}) };
      else if (a === 'setGithubToken') out = setGithubToken_(req.token);
      else if (a === 'githubStatus') out = { ok: true, hasToken: !!PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN') };
      else if (a === 'uploadSong') out = uploadSong_(req.name, req.data);
      else throw new Error('Unknown action');
    }
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- Sheet helpers ---------- */
const SHEET_ID = '1eoV-iLnx5hI1MUF8Dl0EASvu5TZmMUKSQjS7C_By_Pk'; // Wedding Portal DB
let SS_;
function sheet_(name) { SS_ = SS_ || SpreadsheetApp.openById(SHEET_ID); return SS_.getSheetByName(name); }

function ensureHeader_() {
  const sh = sheet_('Guests');
  const head = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  COLS.forEach((c, i) => { if (head[i] !== c) sh.getRange(1, i + 1).setValue(c); });
}

function dateStr_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Colombo', 'yyyy-MM-dd');
  return String(v).slice(0, 10);
}

function readGuests_() {
  ensureHeader_();
  const sh = sheet_('Guests');
  const vals = sh.getDataRange().getValues();
  const head = vals.shift();
  return vals.filter(r => r[0] !== '').map(r => {
    const o = {};
    head.forEach((h, i) => o[h] = r[i] instanceof Date ? r[i].toISOString() : r[i]);
    o.rsvpBy = dateStr_(r[head.indexOf('rsvpBy')]);
    return o;
  });
}

function findRow_(code) {
  const sh = sheet_('Guests');
  const codes = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
  for (let i = 0; i < codes.length; i++) if (String(codes[i][0]) === String(code)) return i + 2;
  return -1;
}

function rowToObj_(row) {
  const sh = sheet_('Guests');
  const vals = sh.getRange(row, 1, 1, COLS.length).getValues()[0];
  const o = {}; COLS.forEach((c, i) => o[c] = vals[i] instanceof Date ? vals[i].toISOString() : vals[i]);
  o.rsvpBy = dateStr_(vals[COLS.indexOf('rsvpBy')]);
  return o;
}

function writeObj_(row, o) {
  ensureHeader_();
  sheet_('Guests').getRange(row, 1, 1, COLS.length).setValues([COLS.map(c => o[c] === undefined ? '' : o[c])]);
}

function newCode_() {
  const chars = 'abcdefghjkmnpqrstuvwxyz23456789';
  const existing = new Set(readGuests_().map(g => String(g.code)));
  let c;
  do { c = ''; for (let i = 0; i < 6; i++) c += chars[Math.floor(Math.random() * chars.length)]; } while (existing.has(c));
  return c;
}

function readSettings_() {
  const vals = sheet_('Settings').getDataRange().getValues(); vals.shift();
  const s = {}; vals.forEach(r => { if (r[0]) s[r[0]] = r[1] instanceof Date ? Utilities.formatDate(r[1], 'Asia/Colombo', 'yyyy-MM-dd') : r[1]; });
  return s;
}

function saveSettings_(updates) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = sheet_('Settings');
    const vals = sh.getDataRange().getValues();
    Object.keys(updates).forEach(k => {
      let i = vals.findIndex(r => r[0] === k);
      if (i < 0) { sh.appendRow([k, updates[k]]); vals.push([k, updates[k]]); }
      else sh.getRange(i + 1, 2).setNumberFormat('@').setValue(updates[k]);
    });
    return readSettings_();
  } finally { lock.releaseLock(); }
}

/* ---------- Admin ---------- */
function saveGuest_(g) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    let row = g.code ? findRow_(g.code) : -1;
    const base = row > 0 ? rowToObj_(row) : {};
    const o = Object.assign({}, base, g);
    if (!o.name) throw new Error('Name is required');
    if (!o.code) o.code = newCode_();
    ['w', 'r'].forEach(ev => {
      const invited = o.invitedTo === 'Both' || o.invitedTo === (ev === 'w' ? 'Wedding' : 'Reception');
      if (invited && !o[ev + '_status']) o[ev + '_status'] = 'Pending';
    });
    o.updatedAt = new Date().toISOString();
    if (o.rsvpBy) o.rsvpBy = "'" + dateStr_(o.rsvpBy);
    ensureHeader_();
    if (row < 0) { sheet_('Guests').appendRow(COLS.map(c => o[c] === undefined ? '' : o[c])); }
    else writeObj_(row, o);
    o.rsvpBy = dateStr_(String(o.rsvpBy || '').replace("'", ''));
    return o;
  } finally { lock.releaseLock(); }
}

function patchGuest_(code, patch) {
  const row = findRow_(code); if (row < 0) throw new Error('Guest not found');
  const o = Object.assign(rowToObj_(row), patch, { updatedAt: new Date().toISOString() });
  writeObj_(row, o); return o;
}

function deleteGuest_(code) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const row = findRow_(code); if (row < 0) return false;
    sheet_('Guests').deleteRow(row); return true;
  } finally { lock.releaseLock(); }
}

/* ---------- Reply-by dates ---------- */
function groupDates_(s) { try { return JSON.parse(s.groupDeadlines || '{}'); } catch (e) { return {}; } }
function deadlineFor_(g, s) {
  const d = (g && g.rsvpBy) || (g && groupDates_(s)[g.group]) || dateStr_(s.rsvpBy);
  return d || '';
}
function niceDate_(d) {
  if (!d) return '';
  const p = d.split('-').map(Number);
  return Utilities.formatDate(new Date(p[0], p[1] - 1, p[2], 12), 'Asia/Colombo', 'd MMMM yyyy');
}
function isClosed_(d) {
  if (!d) return false;
  const today = Utilities.formatDate(new Date(), 'Asia/Colombo', 'yyyy-MM-dd');
  return today > d;
}

/* ---------- Public (invitation page) ---------- */
function guestInfo_(code) {
  const s = readSettings_();
  const row = code ? findRow_(code) : -1;
  const g = row > 0 ? rowToObj_(row) : null;
  const d = deadlineFor_(g, s);
  const settings = { rsvpBy: d, rsvpByText: niceDate_(d), closed: isClosed_(d), song: s.song || '' };
  if (!g) return { ok: true, guest: null, settings };
  const replied = g.w_status === 'Attending' || g.w_status === 'Declined';
  return { ok: true, settings, guest: {
    name: g.name, max: Number(g.max) || 1, invitedTo: g.invitedTo,
    wedding: replied ? { status: g.w_status, veg: Number(g.w_veg) || 0, nonveg: Number(g.w_nonveg) || 0, message: g.w_message } : null
  }};
}

function rsvp_(req) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const s = readSettings_();
    const ev = req.event === 'reception' ? 'r' : 'w';
    const yes = req.attending === 'yes';
    const veg = yes ? Math.max(0, parseInt(req.veg, 10) || 0) : 0;
    const nv = yes ? Math.max(0, parseInt(req.nonveg, 10) || 0) : 0;
    const msg = String(req.message || '').slice(0, 500);
    const patch = {};
    patch[ev + '_status'] = yes ? 'Attending' : 'Declined';
    patch[ev + '_veg'] = veg; patch[ev + '_nonveg'] = nv;
    patch[ev + '_message'] = msg; patch[ev + '_source'] = 'guest';
    patch[ev + '_at'] = new Date().toISOString();

    const row = req.code ? findRow_(req.code) : -1;
    if (row > 0) {
      const g = rowToObj_(row);
      if (isClosed_(deadlineFor_(g, s))) return { ok: false, closed: true, error: 'Replies have closed' };
      const max = Number(g.max) || 1;
      if (yes && veg + nv === 0) throw new Error('Add the number of meals');
      if (veg + nv > max) throw new Error('If more of you are coming, please call us');
      Object.assign(g, patch, { updatedAt: new Date().toISOString() });
      if (g.rsvpBy) g.rsvpBy = "'" + g.rsvpBy;
      writeObj_(row, g);
    } else {
      if (isClosed_(deadlineFor_(null, s))) return { ok: false, closed: true, error: 'Replies have closed' };
      // Reply from a link without a guest code: keep it as a new, unlisted guest
      const name = String(req.name || '').trim().slice(0, 80);
      if (!name) throw new Error('Name is required');
      const o = Object.assign({ code: newCode_(), name: name, side: 'Unlisted', group: 'Replied without a personal link',
        invitedTo: ev === 'w' ? 'Wedding' : 'Reception', max: veg + nv || 1, updatedAt: new Date().toISOString() }, patch);
      ensureHeader_();
      sheet_('Guests').appendRow(COLS.map(c => o[c] === undefined ? '' : o[c]));
    }
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* ---------- Songs: upload to GitHub ---------- */
function setGithubToken_(token) {
  token = String(token || '').trim();
  if (!token) throw new Error('Paste the GitHub key');
  const r = UrlFetchApp.fetch('https://api.github.com/repos/' + REPO, {
    headers: { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json' }, muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error("That key can't open the repository (" + r.getResponseCode() + ')');
  const perms = JSON.parse(r.getContentText()).permissions || {};
  if (!perms.push) throw new Error("That key can read but not upload. Give it Contents: Read and write.");
  PropertiesService.getScriptProperties().setProperty('GITHUB_TOKEN', token);
  return { ok: true, hasToken: true };
}

function uploadSong_(name, b64) {
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) throw new Error('Add the GitHub key in Settings first');
  let file = String(name || '').replace(/\.[^.]+$/, '').replace(/[^A-Za-z0-9 _-]/g, '').trim().replace(/\s+/g, '-');
  if (!file) file = 'song';
  file += '.mp3';
  if (!b64 || b64.length > 14 * 1024 * 1024) throw new Error('Choose an MP3 under 10 MB');
  const url = 'https://api.github.com/repos/' + REPO + '/contents/music/' + encodeURIComponent(file);
  const h = { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json' };
  const ex = UrlFetchApp.fetch(url, { headers: h, muteHttpExceptions: true });
  const body = { message: 'Add song ' + file + ' (from portal)', content: b64 };
  if (ex.getResponseCode() === 200) body.sha = JSON.parse(ex.getContentText()).sha;
  const r = UrlFetchApp.fetch(url, { method: 'put', headers: h, contentType: 'application/json', payload: JSON.stringify(body), muteHttpExceptions: true });
  if (r.getResponseCode() >= 300) throw new Error('GitHub refused the upload (' + r.getResponseCode() + ')');
  return { ok: true, file: file };
}
