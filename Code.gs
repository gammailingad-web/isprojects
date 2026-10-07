/* IS PROJECTS - ISRF backend v2. After pasting: run repair() once, then Deploy > Manage deployments > New version. Login: admin / admin123 */
const SS = SpreadsheetApp.getActiveSpreadsheet();
const H = ['ID','Date','Area','Branch','ErrorType','DocsDamage','ProblemDetails','RequestedBy','ApprovedBy','Decision','ForwardTo','Notes','AnsweredBy','ReviewedBy','Status','CreatedBy','Monitoring','Data'];
const UH = ['Username','Name','Role','Hash','Active','Branch','Dept','Email','Phone','MustChange'], CH = ['Type','Value','Area'];

function sh(n, h) { if (!n || !h) throw new Error('Do not run sh() directly.'); let s = SS.getSheetByName(n); if (!s) { s = SS.insertSheet(n); s.appendRow(h); s.setFrozenRows(1); } return s; }
function hash(p) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'isrf:' + p).map(b => ('0' + (b & 255).toString(16)).slice(-2)).join(''); }
function setup() {
  sh('ISRF', H).getRange('B:B').setNumberFormat('@');
  const u = sh('Users', UH); u.getRange('A:D').setNumberFormat('@'); u.getRange('I:I').setNumberFormat('@');
  if (u.getLastRow() < 2) u.appendRow(['admin','Administrator','Admin',hash('admin123'),true,'','','','','']);
  const c = sh('Config', CH);
  if (c.getLastRow() < 2) {
    [['Area 1','Main Branch'],['Area 1','North Branch'],['Area 2','South Branch']].forEach(b => c.appendRow(['Branch', b[1], b[0]]));
    ['Wrong Encoding','Missing Document','System Error','Duplicate Entry','Wrong Amount'].forEach(e => c.appendRow(['ErrorType', e, '']));
    ['Accounting','CCD'].forEach(d => c.appendRow(['Dept', d, '']));
    c.appendRow(['Monitor', 'ISRF', 'Information System Request Form']);
  }
}
function resetAdmin() {
  setup(); const u = SS.getSheetByName('Users'), ex = rows('Users').find(x => String(x.Username).toLowerCase() === 'admin');
  const row = ['admin','Administrator','Admin',hash('admin123'),true,'','','','',''];
  if (ex) u.getRange(ex._row, 1, 1, 10).setValues([row]); else u.appendRow(row);
}
/* Rebuilds any tab whose header row is outdated (that tab is cleared), then resets admin */
function repair() {
  upgrade();
  [['ISRF', H], ['Users', UH], ['Config', CH]].forEach(([n, h]) => {
    const s = sh(n, h), cur = s.getRange(1, 1, 1, h.length).getValues()[0].map(String);
    if (cur.join('|') !== h.join('|')) { s.clear(); s.appendRow(h); s.setFrozenRows(1); Logger.log('Rebuilt: ' + n); }
  });
  setup(); resetAdmin(); Logger.log('Repair done');
}
/* Adds the Monitoring column to an existing ISRF tab without losing data */
function upg(name, hdr) {
  const s = SS.getSheetByName(name); if (!s) return;
  const cur = s.getRange(1, 1, 1, hdr.length).getValues()[0].map(String);
  for (let i = 0; i < hdr.length; i++) {
    if (cur[i] === hdr[i]) continue;
    if (cur[i] !== '') return;
    s.getRange(1, i + 1).setValue(hdr[i]);
    if (name === 'Users' && hdr[i] === 'Phone') s.getRange('I:I').setNumberFormat('@');
    if (name === 'ISRF' && hdr[i] === 'Monitoring' && s.getLastRow() > 1) s.getRange(2, i + 1, s.getLastRow() - 1, 1).setValue('ISRF');
  }
}
function upgrade() { upg('ISRF', H); upg('Users', UH); upg('Files', FH); }
/* MOVING TO ANOTHER GOOGLE ACCOUNT: run this once in the NEW account after putting the attachment files in a Drive folder named
   "IS PROJECTS Attachments". It reconnects every attachment record to the files in that folder (matching by file name). */
function relinkAttachments() {
  const me = Session.getEffectiveUser().getEmail(), it = DriveApp.getFoldersByName('IS PROJECTS Attachments'); let folder = null;
  while (it.hasNext()) { const f = it.next(); try { if (f.getOwner() && f.getOwner().getEmail() === me) { folder = f; break; } } catch (e) {} }
  if (!folder) throw new Error('Create a Drive folder named "IS PROJECTS Attachments" in this account and put the attachment files in it first');
  cfgSet('FolderId', folder.getId());
  const byName = {}, fi = folder.getFiles(); while (fi.hasNext()) { const x = fi.next(); byName[x.getName().replace(/^Copy of /, '')] = x.getId(); }
  const sheet = fsheet(); let ok = 0, missing = [];
  frows().forEach(r => { const id = byName[r.RecordID + '_' + r.Name]; if (id) { sheet.getRange(r._row, 1).setValue(id); ok++; } else missing.push(r.RecordID + '_' + r.Name); });
  Logger.log('Relinked ' + ok + ' file(s). Not found: ' + missing.length + (missing.length ? ' -> ' + missing.slice(0, 20).join(', ') : ''));
}

function authorize() { DriveApp.getRootFolder(); MailApp.getRemainingDailyQuota(); LockService.getScriptLock().tryLock(1000); LockService.getScriptLock().releaseLock(); CacheService.getScriptCache().put('t', '1', 10); Logger.log('Authorized OK'); }
function testLogin() { Logger.log(JSON.stringify(handle({action:'login', username:'admin', password:'admin123'}))); }

function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function doGet() { return out({ok:true, msg:'IS PROJECTS API'}); }
function doPost(e) {
  let lock = null;
  try { lock = LockService.getScriptLock(); lock.waitLock(20000); return out({ok:true, data: handle(JSON.parse(e.postData.contents))}); }
  catch (x) { return out({ok:false, error: String(x.message || x)}); }
  finally { if (lock) { try { lock.releaseLock(); } catch (y) {} } }
}
function rows(name) {
  const v = SS.getSheetByName(name).getDataRange().getValues(), h = v.shift();
  return v.map((r, i) => { const o = {_row: i + 2}; h.forEach((k, j) => o[k] = r[j] instanceof Date ? Utilities.formatDate(r[j], Session.getScriptTimeZone(), 'yyyy-MM-dd') : r[j]); return o; });
}
function config() {
  const c = {Branch:[], ErrorType:[], Dept:[], Monitor:[], Company:'', Logo:'', UI:{}, Fields:{}};
  rows('Config').forEach(r => {
    if (r.Type === 'Branch') c.Branch.push({name: r.Value, area: r.Area || ''});
    else if (r.Type === 'ErrorType') c.ErrorType.push({name: r.Value, mon: r.Area || ''});
    else if (r.Type === 'Monitor') c.Monitor.push({code: String(r.Value), name: r.Area || String(r.Value)});
    else if (r.Type === 'Fields') { try { c.Fields[r.Value] = JSON.parse(r.Area); } catch (e) {} }
    else if (r.Type === 'Setting') { if (r.Value === 'Company') c.Company = String(r.Area); else if (r.Value === 'Logo') c.Logo = String(r.Area); else if (r.Value === 'UI') { try { c.UI = JSON.parse(r.Area); } catch (e) {} } }
    else if (c[r.Type]) c[r.Type].push(r.Value);
  });
  c.Settings = settings(); { const P = PropertiesService.getScriptProperties(); c.SmsProvider = P.getProperty('SMS_PROVIDER') || 'semaphore'; c.SmsOn = !!(c.SmsProvider === 'android' ? P.getProperty('ANDROID_USER') : P.getProperty('SMS_KEY')); }
  if (!c.Monitor.length) c.Monitor.push({code: 'ISRF', name: 'Information System Request Form'});
  return c;
}
const notes = r => { try { return JSON.parse(r.Notes || '{}'); } catch (e) { return {}; } };
function status(r) {
  if (r.Status === 'Rejected') return 'Rejected';
  if (r.ReviewedBy) return 'Completed';
  if (r.AnsweredBy) return 'For Review';
  if (!r.Decision) return r.ApprovedBy ? 'For ISD Decision' : 'Pending Branch Approval';
  if (r.Decision === 'Execute') return 'For Execution';
  const n = notes(r), p = String(r.ForwardTo).split(',').filter(d => d && !(n[d] && n[d].sig));
  return p.length ? 'Pending Note: ' + p.join(', ') : 'For Execution';
}
const need = (s, roles) => { if (!s || roles.indexOf(s.role) < 0) throw new Error('Not allowed for your role'); };
const sig = s => s.name + '|' + new Date().toISOString();
function save(rec) {
  if (!rec || !rec._row) throw new Error('Invalid ISRF record: spreadsheet row could not be determined.');
  const sheet = SS.getSheetByName('ISRF');
  if (!sheet) throw new Error('ISRF sheet was not found.');
  sheet.getRange(rec._row, 1, 1, H.length).setValues([H.map(k => rec[k] === undefined ? '' : rec[k])]);
}
function find(id) { if (String(id).indexOf('PRF-') === 0) { psheet(); const p = rows('PRF').find(x => x.ID === id); if (!p) throw new Error('Record not found'); p._prf = true; return p; } const r = rows('ISRF').find(x => x.ID === id); if (!r) throw new Error('Record not found'); return r; }
const norm = v => String(v || '').toLowerCase().replace(/\s+/g, ' ').trim();
/* Duplicate = same date + branch + error type + problem details (ignoring case/extra spaces). Rejected requests do not count. */
function findDup(x, skipId) {
  return rows('ISRF').find(r => r.ID !== skipId && r.Status !== 'Rejected' && String(r.Date) === String(x.Date) && norm(r.Monitoring || 'ISRF') === norm(x.Monitoring || 'ISRF') && norm(r.Branch) === norm(x.Branch) &&
    dnorm(r.Data) === dnorm(x.Data) && norm(r.ErrorType) === norm(x.ErrorType) && norm(r.ProblemDetails) === norm(x.ProblemDetails));
}
const dupMsg = r => 'Duplicate: this ISRF already exists as ' + r.ID + ' (' + status(r) + '). Nothing was saved.';
const areaOf = b => { const f = config().Branch.find(x => x.name === b); return f ? f.area : ''; };

/* ---- visibility: Users.Branch column holds the scope: branch names, @AREA, or * (all). Admin always sees everything. ---- */
const scopeOf = s => String(s.branch || '').split(',').map(x => x.trim()).filter(Boolean);
function inScope(s, rec) { if (s.role === 'Admin') return true; const t = scopeOf(s); return t.indexOf('*') >= 0 || t.indexOf(rec.Branch) >= 0 || (!!rec.Area && t.indexOf('@' + rec.Area) >= 0); }
function canSee(s, rec) { return inScope(s, rec) && !(s.role === 'Dept' && String(rec.ForwardTo).split(',').indexOf(s.dept) < 0); }
function findS(s, id) { const r = find(id); if (r._prf) { if (!pvis(s, r, pdef(pdata(r)))) throw new Error('This request is outside your assigned branches'); return r; } if (!canSee(s, r)) throw new Error('This request is outside your assigned branches'); return r; }

/* ---- attachments (stored in one Google Drive folder, metadata in the Files tab) ---- */
const SET = {MaxFileMB: 2, MaxFiles: 5, CapMB: 3000, RetainDays: 365}, FH = ['FileId','RecordID','Name','Size','Mime','UploadedBy','UploadedAt','Ref'];
function settings() { const o = Object.assign({}, SET); rows('Config').forEach(r => { if (r.Type === 'Setting' && SET.hasOwnProperty(r.Value) && r.Area !== '') o[r.Value] = Number(r.Area); }); return o; }
function cfgGet(k) { const f = rows('Config').find(x => x.Type === 'Setting' && x.Value === k); return f ? f.Area : ''; }
function cfgSet(k, v) { const f = rows('Config').find(x => x.Type === 'Setting' && x.Value === k), c = SS.getSheetByName('Config'); if (f) c.getRange(f._row, 3).setValue(v); else c.appendRow(['Setting', k, v]); }
function fsheet() { const had = SS.getSheetByName('Files'), s = sh('Files', FH); if (!had) s.getRange('G:G').setNumberFormat('@'); return s; }
const frows = () => { fsheet(); return rows('Files'); };
function folder() {
  const id = cfgGet('FolderId'); if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder('IS PROJECTS Attachments'); cfgSet('FolderId', f.getId()); return f;
}
function purgeFiles(pred) {
  const sheet = fsheet(), list = frows().filter(pred).sort((a, b) => b._row - a._row);
  list.forEach(f => { try { DriveApp.getFileById(f.FileId).setTrashed(true); } catch (e) {} sheet.deleteRow(f._row); });
  return {n: list.length, bytes: list.reduce((a, f) => a + Number(f.Size || 0), 0)};
}
function oldFiles() {
  const keep = settings().RetainDays; if (!keep) return () => false;
  const recs = {}; rows('ISRF').forEach(x => recs[x.ID] = x);
  return f => { const x = recs[f.RecordID]; if (!x) return true; if (status(x) !== 'Completed') return false;
    const t = new Date(String(x.ReviewedBy).split('|')[1]); return (Date.now() - t.getTime()) / 86400000 > keep; };
}

/* ---- sessions: signed tokens, valid for SESSION_DAYS unless the user signs out, is disabled, or changes password ---- */
const SESSION_DAYS = 30;
function secret() { const P = PropertiesService.getScriptProperties(); let k = P.getProperty('SECRET'); if (!k) { k = Utilities.getUuid() + Utilities.getUuid(); P.setProperty('SECRET', k); } return k; }
function sign(p) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(p, secret())); }
function makeToken(u) { const p = Utilities.base64EncodeWebSafe(JSON.stringify({u: String(u.Username), h: String(u.Hash).slice(0, 10), e: Date.now() + SESSION_DAYS * 86400000})); return p + '.' + sign(p); }
function readToken(t) {
  const bad = new Error('Session expired. Please sign in again.'), parts = String(t || '').split('.');
  if (parts.length !== 2 || sign(parts[0]) !== parts[1]) throw bad;
  const d = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString()); if (d.e < Date.now()) throw bad;
  const u = rows('Users').find(x => String(x.Username) === d.u);
  if (!u || String(u.Hash).slice(0, 10) !== d.h || String(u.Active).toUpperCase() === 'FALSE') throw bad;
  return {username: u.Username, name: u.Name, role: u.Role, branch: u.Branch || '', dept: u.Dept || '', mustChange: String(u.MustChange) === 'Y'};
}
function imgGet(k) { isheet(); return rows('Images').filter(x => x.Key === k).sort((a, b) => a.Idx - b.Idx).map(x => x.Chunk).join(''); }
function setImg(k, v) {
  const sheet = isheet(); rows('Images').filter(x => x.Key === k).sort((a, b) => b._row - a._row).forEach(x => sheet.deleteRow(x._row));
  v = String(v || ''); if (!v) return;
  if (v.indexOf('data:image/') !== 0 || v.length > 240000) throw new Error('Image is too large. Use a smaller picture.');
  for (let i = 0; i * 45000 < v.length; i++) sheet.appendRow([k, i, v.slice(i * 45000, (i + 1) * 45000)]);
}
function isheet() { const had = SS.getSheetByName('Images'), s = sh('Images', ['Key','Idx','Chunk']); if (!had) s.getRange('C:C').setNumberFormat('@'); return s; }
function dnorm(v) { let o = v; if (typeof v === 'string') { try { o = JSON.parse(v || '{}'); } catch (e) { o = {}; } } o = o || {}; return norm(JSON.stringify(Object.keys(o).sort().map(k => [k, norm(o[k])]))); }
function fieldRows(rec, cf, row, E) {
  const fl = cf.Fields[rec.Monitoring] || [{l: 'Error type', c: 'ErrorType'}, {l: 'Documents damaged', c: 'DocsDamage'}, {l: 'Problem details', c: 'ProblemDetails'}];
  let d = {}; try { d = JSON.parse(rec.Data || '{}'); } catch (e) {}
  return fl.map(f => row(E(f.l), E(f.c ? rec[f.c] : d[f.l]).replace(/\n/g, '<br>'))).join('');
}

function mailHtml(rec, msg, cf, s, incRev) {
  const N = notes(rec), E = x => String(x == null ? '' : x).replace(/[&<>]/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[ch]));
  const fm = v => { if (!v) return 'Pending'; const p = String(v).split('|'); return E(p[0]) + ' - ' + Utilities.formatDate(new Date(p[1]), Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm'); };
  const m = cf.Monitor.find(x => x.code === rec.Monitoring) || {name: rec.Monitoring};
  const row = (k, v) => '<tr><td style="padding:7px 10px;border:1px solid #dde4ec;background:#f6f8fa;width:32%"><b>' + k + '</b></td><td style="padding:7px 10px;border:1px solid #dde4ec">' + v + '</td></tr>';
  return '<div style="font-family:Arial,sans-serif;max-width:640px;color:#10243a"><h2 style="margin:0">' + E(cf.Company || 'IS PROJECTS') + '</h2><div style="color:#64748b;margin-bottom:12px">' + E(m.name) + ' - ' + E(rec.ID) + '</div>'
    + (msg ? '<p style="white-space:pre-wrap">' + E(msg) + '</p>' : '') + '<table style="border-collapse:collapse;width:100%">'
    + row('Status', '<b>' + E(status(rec)) + '</b>') + row('Date', E(rec.Date)) + row('Area / Branch', E(rec.Area) + ' / ' + E(rec.Branch)) + fieldRows(rec, cf, row, E) + row('Requested by', fm(rec.RequestedBy)) + row('Approved by (branch)', fm(rec.ApprovedBy))
    + Object.keys(N).map(k => row(E(k) + ' remarks', E(N[k].text).replace(/\n/g, '<br>') + '<br><small>' + fm(N[k].sig) + '</small>')).join('')
    + row('Answered by (ISD)', fm(rec.AnsweredBy)) + (incRev ? row('Reviewed by', fm(rec.ReviewedBy)) : '') + '</table><p style="color:#94a3b8;font-size:12px">Sent by ' + E(s.name) + ' through IS PROJECTS</p></div>';
}

/* ---- messages / follow-ups ---- */
const MH = ['ID','RecordID','At','By','Role','Kind','Text'];
function msheet() { const had = SS.getSheetByName('Messages'), s = sh('Messages', MH); if (!had) s.getRange('C:G').setNumberFormat('@'); return s; }
function addMsg(rid, s, kind, text) { msheet().appendRow(['M' + Date.now().toString(36) + Math.floor(Math.random() * 1000), rid, new Date().toISOString(), s.name, s.role, kind, String(text || '').slice(0, 1000)]); }
function recipients(rec, s) {
  const st = status(rec), us = rows('Users').filter(u => u.Email && String(u.Active).toUpperCase() !== 'FALSE' && u.Username !== s.username);
  const out = us.filter(u => inScope({role: u.Role, branch: u.Branch}, rec) && ((st === 'Pending Branch Approval' && u.Role === 'Approver') || ((st === 'For ISD Decision' || st === 'For Execution') && u.Role === 'ISD') ||
    (st.indexOf('Pending Note') === 0 && u.Role === 'Dept' && st.indexOf(u.Dept) >= 0) || (st === 'For Review' && u.Role === 'Reviewer')));
  if (s.username !== rec.CreatedBy) { const q = us.find(u => u.Username === rec.CreatedBy); if (q && out.indexOf(q) < 0) out.push(q); }
  return out;
}
function notifyMsg(rec, s, kind, text, link) {
  const to = recipients(rec, s).map(u => u.Email); if (!to.length) return 0;
  const cf = config(), E = x => String(x == null ? '' : x).replace(/[&<>]/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[ch]));
  try {
    MailApp.sendEmail({to: to.join(','), name: cf.Company || 'IS PROJECTS', subject: '[' + rec.ID + '] ' + (kind === 'followup' ? 'Follow-up requested' : 'New message'),
      body: s.name + ': ' + (text || 'Asking for a progress update.') + '\nRequest ' + rec.ID + ' - ' + status(rec),
      htmlBody: '<div style="font-family:Arial,sans-serif;max-width:560px;color:#10243a"><h3 style="margin:0">' + E(cf.Company || 'IS PROJECTS') + '</h3><p><b>' + E(s.name) + '</b> (' + E(s.role) + ') ' + (kind === 'followup' ? 'is asking for a progress update on' : 'sent a message about') + ' <b>' + E(rec.ID) + '</b>.</p><p>Status: <b>' + E(status(rec)) + '</b><br>Branch: ' + E(rec.Branch) + '</p>'
        + (text ? '<blockquote style="border-left:3px solid #0f766e;margin:0;padding:6px 12px;background:#f6f8fa">' + E(text).replace(/\n/g, '<br>') + '</blockquote>' : '') + (/^https:\/\//.test(link || '') ? '<p><a href="' + E(link) + '">Open the system</a></p>' : '') + '</div>'});
  } catch (e) { return 0; }
  return to.length;
}

/* ---- account requests (sign-up) ---- */
const SGH = ['ID','At','Username','Name','Email','Phone','Hash','Branch','Note','Status','Kind','ExistingUser','Reason'];
function ssheet() { const had = SS.getSheetByName('Signups'), s = sh('Signups', SGH); if (!had) s.getRange('B:I').setNumberFormat('@'); return s; }
function mailTo(to, subject, html) { if (!to) return false; try { MailApp.sendEmail({to, subject, htmlBody: html, body: html.replace(/<[^>]+>/g, ' '), name: config().Company || 'IS PROJECTS'}); return true; } catch (e) { return false; } }
function notifyAdmins(subject, html) {
  const to = rows('Users').filter(u => u.Role === 'Admin' && u.Email && String(u.Active).toUpperCase() !== 'FALSE').map(u => u.Email); if (!to.length) return 0;
  return mailTo(to.join(','), subject, html) ? to.length : 0;
}
const EH = x => String(x == null ? '' : x).replace(/[&<>]/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[ch]));
function e164(n) { n = String(n).replace(/[^\d+]/g, ''); if (/^0\d{10}$/.test(n)) return '+63' + n.slice(1); if (/^63\d{10}$/.test(n)) return '+' + n; return n; }
function gatewaySms(to, message) {
  const P = PropertiesService.getScriptProperties(), prov = P.getProperty('SMS_PROVIDER') || 'semaphore';
  const nums = String(to || '').split(/[,;\s]+/).filter(x => /^\+?\d{10,13}$/.test(x)); if (!nums.length || nums.length > 5) throw new Error('Enter 1 to 5 valid mobile numbers');
  const text = String(message || '').slice(0, 459);
  if (prov === 'android') {
    const u = P.getProperty('ANDROID_USER'), p = P.getProperty('ANDROID_PASS'); if (!u || !p) throw new Error('SMS gateway is not configured');
    const res = UrlFetchApp.fetch('https://api.sms-gate.app/3rdparty/v1/messages', {method: 'post', contentType: 'application/json', headers: {Authorization: 'Basic ' + Utilities.base64Encode(u + ':' + p)}, payload: JSON.stringify({message: text, phoneNumbers: nums.map(e164)}), muteHttpExceptions: true});
    if (res.getResponseCode() >= 300) throw new Error('SMS gateway error ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 120)); return nums.length;
  }
  const key = P.getProperty('SMS_KEY'); if (!key) throw new Error('SMS gateway is not configured');
  const pl = {apikey: key, number: nums.join(','), message: text}; if (P.getProperty('SMS_SENDER')) pl.sendername = P.getProperty('SMS_SENDER');
  const res = UrlFetchApp.fetch('https://api.semaphore.co/api/v4/messages', {method: 'post', payload: pl, muteHttpExceptions: true});
  if (res.getResponseCode() >= 300) throw new Error('SMS gateway error: ' + res.getContentText().slice(0, 120)); return nums.length;
}

/* ---- database export / transfer to SQL Server ---- */
const SQLT = {ISRF: {pk: 'ID'}, PRF: {pk: 'ID'}, Users: {pk: 'Username', bit: ['Active']}, Messages: {pk: 'ID'}, Signups: {pk: 'ID'}, Files: {pk: 'FileId'}, Config: {}};
function exportAll() {
  return ['ISRF','PRF','Users','Config','Messages','Signups','Files'].filter(n => SS.getSheetByName(n)).map(n => {
    const v = SS.getSheetByName(n).getDataRange().getValues(), cols = v.shift().map(String), t = SQLT[n] || {};
    const out = v.filter(r => r.join('') !== '').map(r => r.map((x, j) => x instanceof Date ? Utilities.formatDate(x, Session.getScriptTimeZone(), 'yyyy-MM-dd') : (n === 'Signups' && cols[j] === 'Hash') ? '' : (x === true || x === false) ? x : String(x == null ? '' : x)));
    return {name: n, cols, rows: out, pk: t.pk || null, bit: t.bit || []};
  });
}
function transferSql(p) {
  const host = String(p.host || '').trim(), db = String(p.db || '').trim();
  if (!host || !db || !p.user) throw new Error('Host, database and user are required');
  if (!/^[A-Za-z0-9._\-]+$/.test(host) || !/^[A-Za-z0-9_\-]+$/.test(db)) throw new Error('Invalid host or database name');
  const conn = Jdbc.getConnection('jdbc:sqlserver://' + host + ':' + (Number(p.port) || 1433) + ';databaseName=' + db + ';encrypt=true;trustServerCertificate=' + (p.trust ? 'true' : 'false'), p.user, p.pass), out = {};
  conn.setAutoCommit(false);
  try {
    const st = conn.createStatement();
    exportAll().forEach(t => {
      const defs = t.cols.map(c => '[' + c + '] ' + (c === t.pk ? 'NVARCHAR(100) NOT NULL PRIMARY KEY' : t.bit.indexOf(c) >= 0 ? 'BIT NULL' : 'NVARCHAR(MAX) NULL')).join(', ');
      st.execute("IF OBJECT_ID(N'dbo." + t.name + "', N'U') IS NULL CREATE TABLE [dbo].[" + t.name + '] (' + defs + ')');
      if (p.replace) st.execute('DELETE FROM [dbo].[' + t.name + ']');
      out[t.name] = t.rows.length; if (!t.rows.length) return;
      const ps = conn.prepareStatement('INSERT INTO [dbo].[' + t.name + '] (' + t.cols.map(c => '[' + c + ']').join(', ') + ') VALUES (' + t.cols.map(() => '?').join(', ') + ')');
      t.rows.forEach((r, k) => { r.forEach((v, j) => { if (t.bit.indexOf(t.cols[j]) >= 0) ps.setBoolean(j + 1, v === true || String(v).toUpperCase() === 'TRUE'); else ps.setString(j + 1, String(v)); }); ps.addBatch(); if ((k + 1) % 200 === 0) ps.executeBatch(); });
      ps.executeBatch(); ps.close();
    });
    conn.commit();
  } catch (e) { try { conn.rollback(); } catch (x) {} throw e; } finally { conn.close(); }
  return out;
}

/* ---- Procurement Request Form (PRF) ---- */
const PH = ['ID','Date','Area','Branch','Reason','Status','CreatedBy','Data'];
const PRF_GROUPS = {'SYSTEM UNIT': ['Motherboard','Memory (RAM)','HDD/SSD','Power supply','Power cord','CPU fan'], 'CCTV': ['Siamese wire or connector','Video power supply','DVR power supply','Hard disk drive','Camera','DVR'],
  'PRINTER DOT MATRIX': ['Printer cord','Power cord','Ribbon','Printer head'], 'PRINTER INKJET': ['Printer cord','Power cord','Ink','Printer head','Roller feeder','Scanner head'], 'UPS': ['Battery','Power button','Fuse','Power cord','Main board'],
  'MONITOR': ['Screen','Power cord/power supply','VGA/HDMI cable','VGA/HDMI port','Main board'], 'SWITCH HUB': ['Power cord','Port','Power'], 'LAPTOP': ['Screen','Charger','Keyboard','Battery','HDD/SSD','RAM','Motherboard'], 'OTHER': ['Headset/headphone']};
const PRF_FLOW = {name:'PRF', steps: [{k:'AM', l:'Area Manager', u:[]}, {k:'ISM', l:'Information System Manager', u:[]}, {k:'COO', l:'Chief Operating Officer', u:[]}], acct: 'Accounting'};
function psheet() { const had = SS.getSheetByName('PRF'), s = sh('PRF', PH); if (!had) s.getRange('B:B').setNumberFormat('@'); return s; }
function prfCfg() { let g, f; try { g = JSON.parse(cfgGet('PRFGroups')); } catch (e) {} try { f = JSON.parse(cfgGet('PRFFlow')); } catch (e) {} return {groups: g && Object.keys(g).length ? g : PRF_GROUPS, flow: f && f.steps ? f : PRF_FLOW}; }
function pdata(rec) { try { return JSON.parse(rec.Data || '{}'); } catch (e) { return {}; } }
function pdef(d) { d.items = d.items || []; d.diag = d.diag || {items: {}, remarks: ''}; d.sign = d.sign || {}; d.sign.approvals = d.sign.approvals || {}; d.flow = d.flow || {steps: []}; d.rem = d.rem || {}; d.canvass = d.canvass || []; d.award = d.award || {}; d.buy = d.buy || {}; d.dates = d.dates || {}; d.audit = d.audit || {}; return d; }
function pendingApprovals(d) { const o = []; if (!d.sign.verifier) o.push('Accounting'); d.flow.steps.forEach(k => { if (!d.sign.approvals[k]) o.push(k); }); return o; }
function pstat(rec, d) {
  const s = d.sign;
  if (d.rejected) return 'Rejected'; if (s.checker) return 'Completed'; if (s.notation) return 'For Check';
  if (s.purchaser) return d.dates.received ? 'For Notation' : 'For Receiving';
  if (!s.branch && !s.assessor) return 'Pending Branch Approval';
  if (!s.assessor) return 'For ISD Assessment'; if (!d.flow.forwarded) return 'For Forwarding';
  // Accounting is a dedicated stage in the request list. Keep it separate
  // from the later configured approval steps so users can immediately see
  // that the request is waiting for Accounting review/verification.
  if (!s.verifier) return 'Pending for Accounting';
  const p = pendingApprovals(d); if (p.length) return 'Pending: ' + p.join(', ');
  const aw = d.items.filter((x, i) => d.award[i]).length; if (aw < d.items.length) return d.canvass.length ? 'For Award' : 'For Canvass';
  return 'For Purchase';
}
function pvis(s, rec, d) {
  if (!inScope(s, rec)) return false;
  if (s.role === 'Admin') return true;
  if (s.role === 'User') return rec.CreatedBy === s.username;
  const me = String(s.username || '').toLowerCase();
  const fl = prfCfg().flow || PRF_FLOW;
  const assignedStep = Array.isArray(d.flow && d.flow.steps) && d.flow.steps.some(k => {
    const st = (fl.steps || []).find(x => x.k === k);
    return st && Array.isArray(st.u) && st.u.some(u => String(u).toLowerCase() === me);
  });
  const diagUser = Array.isArray(fl.diag) && fl.diag.length && fl.diag.some(u => String(u).toLowerCase() === me);
  const awardUser = Array.isArray(fl.award) && fl.award.length && fl.award.some(u => String(u).toLowerCase() === me);
  const acctUser = s.role === 'Dept' && String(s.dept || '').toLowerCase() === String(fl.acct || 'Accounting').toLowerCase();
  // ISD/Admin must be able to see a PRF at any stage so ISD notation,
  // diagnostic, canvass and other authorized actions can be performed.
  if (s.role === 'ISD' || s.role === 'Reviewer' || diagUser || awardUser || acctUser) return true;
  // Configured approval users can see the request when it has reached the
  // approval flow, but only their own configured step can be signed.
  if (assignedStep) return !!(d.flow && d.flow.forwarded);
  // Branch Manager/Approver can see pending branch approvals in their scope.
  if (s.role === 'Approver') return !d.sign.branch;
  return false;
}
function prfNotify(rec, role, subject) {
  const to = rows('Users').filter(u => u.Role === role && u.Email && String(u.Active).toUpperCase() !== 'FALSE' && inScope({role: u.Role, branch: u.Branch}, rec)).map(u => u.Email);
  if (to.length) mailTo(to.join(','), subject, '<div style="font-family:Arial,sans-serif;max-width:560px;color:#10243a"><p><b>' + EH(rec.ID) + '</b> - ' + EH(rec.Branch) + '</p><p>' + EH(subject) + '</p><p>Open the Procurement page to review it.</p></div>');
}
function psave(rec, d) {
  const j = JSON.stringify(d); if (j.length > 49000) throw new Error('This request has too much data. Remove some canvass entries.');
  rec.Data = j; rec.Status = pstat(rec, d); psheet().getRange(rec._row, 1, 1, PH.length).setValues([PH.map(k => rec[k] === undefined ? '' : rec[k])]);
}
function cleanItems(a) {
  if (!Array.isArray(a) || !a.length) throw new Error('Add at least one item'); if (a.length > 30) throw new Error('Up to 30 items per request');
  return a.map(x => {
    const n = String(x.n || x.name || x.description || '').trim(), q = Number(x.q ?? x.qty), p = Number(x.p ?? x.price);
    if (!n) throw new Error('Every row needs an item description');
    if (!(q > 0)) throw new Error('Quantity must be more than 0');
    if (!(p >= 0)) throw new Error('Enter the projected amount');
    const c = (x.c === '' || x.c == null) ? null : Number(x.c);
    if (c !== null && !(c >= 0)) throw new Error('Enter a valid account charge');
    return {
      n: n.slice(0, 200),
      q,
      p,
      c,
      unit: String(x.unit || '').trim().slice(0, 40),
      group: String(x.group || '').trim().slice(0, 80),
      remarks: String(x.remarks || '').trim().slice(0, 300)
    };
  });
}

function verifyCredential(username, password) {
  const un = String(username || '').trim();
  const pw = String(password || '');
  if (!un || !pw) throw new Error('Username and password are required.');
  const u = rows('Users').find(x => String(x.Username).toLowerCase() === un.toLowerCase());
  if (!u || String(u.Hash).trim() !== hash(pw) || String(u.Active).toUpperCase() === 'FALSE') throw new Error('Wrong username or password.');
  return {username: u.Username, name: u.Name, role: u.Role, branch: u.Branch || '', dept: u.Dept || '', mustChange: String(u.MustChange) === 'Y'};
}

function handle(r) {
  const cache = CacheService.getScriptCache();
  if (r.action === 'login') {
    const u = rows('Users').find(x => String(x.Username).toLowerCase() === String(r.username).toLowerCase());
    if (!u || String(u.Hash).trim() !== hash(String(r.password)) || String(u.Active).toUpperCase() === 'FALSE') throw new Error('Wrong username or password');
    const s = {username: u.Username, name: u.Name, role: u.Role, branch: u.Branch || '', dept: u.Dept || '', mustChange: String(u.MustChange) === 'Y'}; return {token: makeToken(u), user: s};
  }
  if (r.action === 'brand') {
    const b = config(); let L = {}; try { L = JSON.parse(cfgGet('Login') || '{}'); } catch (e) {}
    return {Company: b.Company, Logo: b.Logo, UI: b.UI, Login: L, LoginImg: imgGet('login'), Branches: b.Branch.map(x => x.name)};
  }
  if (r.action === 'forgot') {
    const k = String(r.id || '').trim().toLowerCase(), msg = 'If an account matches, the administrator has been notified and will send you a temporary password by email or text.';
    if (!k || k.length > 80) throw new Error('Enter your username or email');
    if (cache.get('fg_' + k)) return msg; cache.put('fg_' + k, '1', 60);
    const u = rows('Users').find(x => String(x.Username).toLowerCase() === k || (x.Email && String(x.Email).toLowerCase() === k));
    if (u && String(u.Active).toUpperCase() !== 'FALSE') {
      ssheet();
      if (!rows('Signups').some(x => x.Status === 'Pending' && x.Kind === 'reset' && x.ExistingUser === u.Username)) {
        SS.getSheetByName('Signups').appendRow(['S' + Date.now().toString(36), new Date().toISOString(), u.Username, u.Name, u.Email || '', u.Phone || '', '', '', '', 'Pending', 'reset', u.Username, '']);
        notifyAdmins('Password reset requested: ' + u.Username, '<div style="font-family:Arial,sans-serif;max-width:560px;color:#10243a"><h3>Password reset request</h3><p><b>' + EH(u.Name) + '</b> (' + EH(u.Username) + ') forgot the password.<br>Open Setup and users > Account requests and choose Send temporary password.</p></div>');
      }
    }
    return msg;
  }
  if (r.action === 'signup') {
    const d = r.data || {}; let L = {}; try { L = JSON.parse(cfgGet('Login') || '{}'); } catch (e) {}
    if (L.signup === false) throw new Error('Account requests are turned off. Please contact the administrator.');
    const msg = 'Request received. You will be notified by email or text once the administrator approves your account. If an account already exists for these details, the administrator is notified so it can be updated.';
    if (d.website) return msg;
    const un = String(d.username || '').trim(), nm = String(d.name || '').trim(), em = String(d.email || '').trim().toLowerCase(), ph = String(d.phone || '').replace(/[^\d+]/g, ''), pw = String(d.password || '');
    if (!/^[A-Za-z0-9._-]{3,30}$/.test(un)) throw new Error('Username must be 3 to 30 letters, numbers, dots, dashes or underscores');
    if (nm.length < 3) throw new Error('Enter your full name');
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(em)) throw new Error('Enter a valid email address');
    if (!/^\+?\d{10,13}$/.test(ph)) throw new Error('Enter a valid contact number');
    if (pw.length < 8) throw new Error('Password must be at least 8 characters');
    if (cache.get('su_' + un.toLowerCase())) throw new Error('Please wait a minute before trying again');
    cache.put('su_' + un.toLowerCase(), '1', 60);
    ssheet(); const pend = rows('Signups').filter(x => x.Status === 'Pending');
    if (pend.length >= 60) throw new Error('Too many pending requests. Please contact the administrator.');
    if (pend.some(x => String(x.Username).toLowerCase() === un.toLowerCase() || String(x.Email).toLowerCase() === em)) return msg;
    const ex = rows('Users').find(x => String(x.Username).toLowerCase() === un.toLowerCase() || (x.Email && String(x.Email).toLowerCase() === em));
    SS.getSheetByName('Signups').appendRow(['S' + Date.now().toString(36), new Date().toISOString(), un, nm, em, ph, ex ? '' : hash(pw), String(d.branch || '').slice(0, 120), String(d.note || '').slice(0, 300), 'Pending', ex ? 'existing' : 'new', ex ? ex.Username : '', '']);
    notifyAdmins(ex ? 'Existing account needs an update: ' + un : 'New account request: ' + un,
      '<div style="font-family:Arial,sans-serif;max-width:560px;color:#10243a"><h3>' + (ex ? 'An existing account asked for access or an update' : 'New account request') + '</h3><p><b>' + EH(nm) + '</b> (' + EH(un) + ')<br>Email: ' + EH(em) + '<br>Contact: ' + EH(ph) + '<br>Branch: ' + EH(d.branch || '-') + '<br>Note: ' + EH(d.note || '-') + '</p>'
      + (ex ? '<p>This username or email already belongs to the account <b>' + EH(ex.Username) + '</b>. Review it under Setup and users > Account requests to update or modify the account.</p>' : '<p>Open Setup and users > Account requests to approve or reject it.</p>') + '</div>');
    return msg;
  }
  const s = readToken(r.token), adm = s.role === 'Admin'; upgrade();
  switch (r.action) {
    case 'me': return {user: s};
    case 'list': {
      let rs = rows('ISRF').filter(x => canSee(s, x));
      rs.forEach(x => x.Monitoring = x.Monitoring || 'ISRF');
      return {rows: rs, cfg: config()};
    }
    case 'create': {
      need(s, ['Admin','ISD','User']);
      const d = r.data, sheet = SS.getSheetByName('ISRF'), br = d.Branch;
      if (!inScope(s, {Branch: br, Area: areaOf(br)})) throw new Error('You can only create requests for your assigned branches');
      const cf = config(), mc = cf.Monitor.some(m => m.code === d.Monitoring) ? d.Monitoring : cf.Monitor[0].code;
      (cf.Fields[mc] || []).forEach(f => { if (f.r && !String((f.c ? d[f.c] : (d.Data || {})[f.l]) || '').trim()) throw new Error(f.l + ' is required'); });
      const dup = findDup({Date: d.Date, Branch: br, ErrorType: d.ErrorType, ProblemDetails: d.ProblemDetails, Monitoring: mc, Data: d.Data}); if (dup) throw new Error(dupMsg(dup));
      const pre = mc + '-' + new Date().getFullYear() + '-'; let mx = 0;
      rows('ISRF').forEach(x => { if (String(x.ID).indexOf(pre) === 0) mx = Math.max(mx, parseInt(String(x.ID).slice(pre.length)) || 0); });
      const rec = {ID: pre + ('0000' + (mx + 1)).slice(-4), Monitoring: mc, Date: d.Date, Area: areaOf(br), Branch: br, ErrorType: d.ErrorType,
        DocsDamage: d.DocsDamage, ProblemDetails: d.ProblemDetails, Data: JSON.stringify(d.Data || {}), RequestedBy: sig(s), CreatedBy: s.username};
      rec.Status = status(rec); sheet.appendRow(H.map(k => rec[k] || '')); return rec.ID;
    }
    case 'update': {
      need(s, ['Admin','ISD']); const rec = findS(s, r.id);
      if (rec.AnsweredBy || rec.ReviewedBy || rec.Status === 'Rejected') throw new Error('This request is locked. The administrator must return it first.');
      const dup = findDup(Object.assign({}, r.data, {Monitoring: rec.Monitoring}), r.id); if (dup) throw new Error(dupMsg(dup));
      ['Date','Branch','ErrorType','DocsDamage','ProblemDetails'].forEach(k => rec[k] = r.data[k]); rec.Data = JSON.stringify(r.data.Data || {}); rec.Area = areaOf(rec.Branch); save(rec); return true;
    }
    case 'delete': { need(s, ['Admin','ISD']); const rec = findS(s, r.id); if (status(rec) !== 'Pending Branch Approval') throw new Error('A request can only be deleted while it is pending branch approval'); purgeFiles(f => f.RecordID === rec.ID); msheet(); rows('Messages').filter(m => m.RecordID === rec.ID).sort((a, b) => b._row - a._row).forEach(m => SS.getSheetByName('Messages').deleteRow(m._row)); SS.getSheetByName('ISRF').deleteRow(rec._row); return true; }
    case 'act': {
      const rec = findS(s, r.id), d = r.data || {}, st = status(rec), n = notes(rec);
      if (st === 'Completed' || st === 'Rejected') throw new Error('This request is already closed');
      switch (r.type) {
        case 'approve':
          if (!(adm || s.role === 'Approver')) throw new Error('Only the approver of this branch can approve');
          rec.ApprovedBy = sig(s); break;
        case 'decide':
          need(s, ['Admin','ISD']); if (rec.Decision) throw new Error('Decision already made');
          if (d.decision === 'Forward') {
            if (!rec.ApprovedBy) throw new Error('Branch approval is required before forwarding');
            if (!d.forward || !d.forward.length) throw new Error('Choose at least one department');
            rec.Decision = 'Forward'; rec.ForwardTo = d.forward.join(',');
          } else rec.Decision = 'Execute';
          break;
        case 'note': {
          if (!(adm || s.role === 'Dept')) throw new Error('Not allowed for your role');
          const dep = adm ? d.dept : s.dept;
          if (rec.Decision !== 'Forward' || String(rec.ForwardTo).split(',').indexOf(dep) < 0) throw new Error('This request was not forwarded to ' + dep);
          n[dep] = {text: d.text || '', sig: sig(s)}; rec.Notes = JSON.stringify(n); break;
        }
        case 'answer':
          need(s, ['Admin','ISD']); if (st !== 'For Execution') throw new Error('Not ready for execution');
          if (d.text) { n.ISD = {text: d.text, sig: sig(s)}; rec.Notes = JSON.stringify(n); }
          rec.AnsweredBy = sig(s); break;
        case 'review': need(s, ['Admin','Reviewer']); if (st !== 'For Review') throw new Error('Not ready for review'); rec.ReviewedBy = sig(s); break;
        case 'reject': need(s, ['Admin','ISD','Approver']); if (rec.AnsweredBy || rec.ReviewedBy) throw new Error('Answered requests can only be returned by the administrator'); rec.Status = 'Rejected'; break;
        default: throw new Error('Unknown step');
      }
      rec.Status = r.type === 'reject' ? 'Rejected' : status(rec); save(rec); return rec.Status;
    }
    case 'files': { findS(s, r.id); return frows().filter(f => f.RecordID === r.id).map(f => ({FileId: f.FileId, Name: f.Name, Size: f.Size, UploadedBy: f.UploadedBy, UploadedAt: f.UploadedAt, Ref: f.Ref || ''})); }
    case 'upload': {
      need(s, ['Admin','ISD','User','Approver','Dept','Reviewer']);
      const rec = findS(s, r.id), st = rec._prf ? rec.Status : status(rec); if (String(r.ref || '').indexOf('cv:') === 0) { const prf = rec._prf ? pdef(pdata(rec)) : null, fl = prfCfg().flow; const can = prf && (adm || rec.CreatedBy.toLowerCase() === s.username.toLowerCase() || s.role === 'ISD' || (Array.isArray(fl.award) && fl.award.some(u => String(u).toLowerCase() === s.username.toLowerCase()))); if (!can) throw new Error('You are not allowed to attach a canvass quotation'); } if (!adm && (st === 'Completed' || st === 'Rejected')) throw new Error('This request is closed');
      const set = settings(), ext = String(r.name).split('.').pop().toLowerCase();
      if (['pdf','jpg','jpeg','png'].indexOf(ext) < 0) throw new Error('Only PDF, JPG and PNG files are allowed');
      const bytes = Utilities.base64Decode(r.b64), fr = frows();
      if (bytes.length > set.MaxFileMB * 1048576) throw new Error('File is larger than ' + set.MaxFileMB + ' MB');
      if (fr.filter(f => f.RecordID === rec.ID).length >= (rec._prf ? set.MaxFiles * 4 : set.MaxFiles)) throw new Error('Maximum ' + set.MaxFiles + ' files per request');
      if (fr.reduce((a, f) => a + Number(f.Size || 0), 0) + bytes.length > set.CapMB * 1048576) throw new Error('Storage limit reached. Ask the administrator to clean up old attachments.');
      const file = folder().createFile(Utilities.newBlob(bytes, r.mime || 'application/octet-stream', rec.ID + '_' + r.name));
      fsheet().appendRow([file.getId(), rec.ID, r.name, bytes.length, r.mime || '', s.username, new Date().toISOString(), r.ref || '']); return true;
    }
    case 'getFile': {
      const f = frows().find(x => x.FileId === r.fileId); if (!f) throw new Error('File not found'); findS(s, f.RecordID);
      return {name: f.Name, mime: f.Mime, b64: Utilities.base64Encode(DriveApp.getFileById(f.FileId).getBlob().getBytes())};
    }
    case 'delFile': {
      const f = frows().find(x => x.FileId === r.fileId); if (!f) throw new Error('File not found'); findS(s, f.RecordID);
      if (!(adm || s.role === 'ISD' || f.UploadedBy === s.username)) throw new Error('Not allowed for your role');
      purgeFiles(x => x.FileId === f.FileId); return true;
    }
    case 'storage': {
      need(s, ['Admin']); const fr = frows(), pred = oldFiles(), set = settings();
      return {used: fr.reduce((a, f) => a + Number(f.Size || 0), 0), count: fr.length, set, eligible: fr.filter(pred).length, mailQuota: MailApp.getRemainingDailyQuota()};
    }
    case 'cleanup': need(s, ['Admin']); return purgeFiles(oldFiles());
    case 'saveSettings': need(s, ['Admin']); Object.keys(SET).forEach(k => { const n = Number(r.data[k]); if (isFinite(n) && n >= 0) cfgSet(k, n); }); return true;
    case 'email': {
      need(s, ['Admin','ISD','User','Approver','Dept','Reviewer']);
      const rec = findS(s, r.id), to = String(r.to || '').split(/[,;\s]+/).filter(Boolean);
      if (!to.length || to.length > 10 || to.some(x => !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x))) throw new Error('Enter 1 to 10 valid email addresses');
      const cf = config(), m = cf.Monitor.find(x => x.code === rec.Monitoring) || {name: rec.Monitoring};
      const msg = {to: to.join(','), subject: '[' + rec.ID + '] ' + m.name + ' - ' + status(rec), body: rec.ID + ' - ' + status(rec) + '\n' + (r.message || ''), htmlBody: mailHtml(rec, r.message, cf, s, adm && r.incReview), name: cf.Company || 'IS PROJECTS'};
      if (r.attach) msg.attachments = frows().filter(f => f.RecordID === rec.ID).slice(0, 5).map(f => DriveApp.getFileById(f.FileId).getBlob().setName(f.Name));
      MailApp.sendEmail(msg); return to.length;
    }
    case 'sms': need(s, ['Admin','ISD','User','Approver','Dept','Reviewer']); findS(s, r.id); return gatewaySms(r.to, r.message);
    case 'saveSms': {
      need(s, ['Admin']); const P = PropertiesService.getScriptProperties();
      if (r.clear) { ['SMS_KEY','SMS_SENDER','SMS_PROVIDER','ANDROID_USER','ANDROID_PASS'].forEach(k => P.deleteProperty(k)); return true; }
      P.setProperty('SMS_PROVIDER', r.provider === 'android' ? 'android' : 'semaphore');
      if (r.key) P.setProperty('SMS_KEY', String(r.key).trim()); if (r.user) P.setProperty('ANDROID_USER', String(r.user).trim()); if (r.pass) P.setProperty('ANDROID_PASS', String(r.pass).trim());
      P.setProperty('SMS_SENDER', String(r.sender || '').trim()); return true;
    }
    case 'saveBrand': {
      need(s, ['Admin']); const d = r.data;
      if (d.logo !== undefined && d.logo !== null) { if (String(d.logo).length > 45000) throw new Error('Logo is too large'); cfgSet('Logo', d.logo); }
      cfgSet('Company', String(d.company || '').slice(0, 120)); cfgSet('UI', JSON.stringify(d.ui || {})); return true;
    }
    case 'home': {
      isheet(); const imgs = {}; rows('Images').sort((a, b) => a.Idx - b.Idx).forEach(x => { if (x.Key !== 'login') imgs[x.Key] = (imgs[x.Key] || '') + x.Chunk; });
      let home = {}; try { home = JSON.parse(cfgGet('Home') || '{}'); } catch (e) {}
      return {home, images: imgs};
    }
    case 'msgs': { findS(s, r.id); msheet(); return rows('Messages').filter(m => m.RecordID === r.id).map(m => ({By: m.By, Role: m.Role, Kind: m.Kind, At: m.At, Text: m.Text})); }
    case 'msg': {
      need(s, ['Admin','ISD','User','Approver','Dept','Reviewer']); const rec = findS(s, r.id), kind = r.kind === 'followup' ? 'followup' : 'message', text = String(r.text || '').trim();
      if (kind === 'message' && !text) throw new Error('Type a message first');
      msheet();
      if (kind === 'followup') { const last = rows('Messages').filter(m => m.RecordID === rec.ID && m.Kind === 'followup' && m.By === s.name).pop(); if (last && Date.now() - new Date(last.At).getTime() < 4 * 3600000) throw new Error('You already asked for an update recently. Please wait a few hours.'); }
      addMsg(rec.ID, s, kind, text || 'Asking for a progress update.'); return notifyMsg(rec, s, kind, text, r.link);
    }
    case 'inbox': {
      msheet(); const vis = {}; rows('ISRF').filter(x => canSee(s, x) && !(s.role === 'User' && x.CreatedBy !== s.username)).forEach(x => vis[x.ID] = x); const cut = Date.now() - 30 * 86400000;
      return rows('Messages').filter(m => vis[m.RecordID] && m.By !== s.name && new Date(m.At).getTime() > cut).slice(-40).reverse().map(m => ({RecordID: m.RecordID, At: m.At, By: m.By, Role: m.Role, Kind: m.Kind, Text: m.Text, Status: status(vis[m.RecordID])}));
    }
    case 'return': {
      need(s, ['Admin']); const rec = findS(s, r.id), d = r.data || {}, n = notes(rec);
      const locked = !!(rec.AnsweredBy || rec.ReviewedBy || rec.Status === 'Rejected');
      if (d.to === 'pending' && locked) throw new Error('Use Return to send answered, reviewed or rejected requests back');
      if (d.to !== 'pending' && !locked) throw new Error('Only answered, reviewed or rejected requests can be returned');
      let label = '';
      if (d.to === 'approval' || d.to === 'pending') { rec.ApprovedBy = ''; rec.Decision = ''; rec.ForwardTo = ''; rec.AnsweredBy = ''; rec.ReviewedBy = ''; Object.keys(n).forEach(k => { delete n[k].sig; }); label = d.to === 'pending' ? 'pending branch approval' : 'branch approval'; }
      else if (d.to === 'notation') {
        const deps = (d.depts && d.depts.length) ? d.depts : String(rec.ForwardTo).split(',').filter(Boolean);
        if (!deps.length) throw new Error('Choose at least one department'); if (!rec.ApprovedBy) throw new Error('Branch approval is required before notation');
        rec.Decision = 'Forward'; rec.ForwardTo = deps.join(','); deps.forEach(k => { if (n[k]) delete n[k].sig; }); rec.AnsweredBy = ''; rec.ReviewedBy = ''; label = 'notation by ' + deps.join(', ');
      } else if (d.to === 'execution') { rec.AnsweredBy = ''; rec.ReviewedBy = ''; if (!rec.Decision) rec.Decision = 'Execute'; label = 'execution'; }
      else throw new Error('Choose where to return the request');
      rec.Notes = JSON.stringify(n); rec.Status = ''; rec.Status = status(rec); save(rec);
      addMsg(rec.ID, s, 'return', 'Returned for ' + label + (d.reason ? ': ' + d.reason : '')); return rec.Status;
    }
    case 'signups': {
      need(s, ['Admin']); ssheet(); const list = rows('Signups').map(x => ({ID: x.ID, At: x.At, Username: x.Username, Name: x.Name, Email: x.Email, Phone: x.Phone, Branch: x.Branch, Note: x.Note, Status: x.Status, Kind: x.Kind, ExistingUser: x.ExistingUser, Reason: x.Reason}));
      return list.filter(x => x.Status === 'Pending').concat(list.filter(x => x.Status !== 'Pending').slice(-10).reverse());
    }
    case 'approveSignup': {
      need(s, ['Admin']); ssheet(); const x = rows('Signups').find(y => y.ID === r.id), d = r.data || {};
      if (!x || x.Status !== 'Pending' || x.Kind !== 'new') throw new Error('This request is no longer pending');
      if (['Admin','ISD','User','Approver','Dept','Reviewer','Viewer'].indexOf(d.role) < 0) throw new Error('Choose a role');
      if (rows('Users').some(u => String(u.Username).toLowerCase() === String(x.Username).toLowerCase())) throw new Error('That username now exists. Dismiss this request and update the existing account.');
      SS.getSheetByName('Users').appendRow([x.Username, x.Name, d.role, x.Hash, true, d.branch || '', d.dept || '', x.Email, x.Phone, '']);
      const sg = SS.getSheetByName('Signups'); sg.getRange(x._row, 7).setValue(''); sg.getRange(x._row, 10).setValue('Approved');
      const cf = config(), link = /^https:\/\//.test(r.link || '') ? r.link : '';
      const emailed = mailTo(x.Email, 'Your account is approved - ' + (cf.Company || 'IS PROJECTS'),
        '<div style="font-family:Arial,sans-serif;max-width:560px;color:#10243a"><h3>' + EH(cf.Company || 'IS PROJECTS') + '</h3><p>Hello ' + EH(x.Name) + ', your account has been approved.</p><p>Username: <b>' + EH(x.Username) + '</b><br>Role: ' + EH(d.role) + '<br>Sign in with the password you chose when you requested the account.</p>' + (link ? '<p><a href="' + EH(link) + '">Open the system</a></p>' : '') + '</div>');
      return {emailed, username: x.Username, name: x.Name, email: x.Email, phone: x.Phone};
    }
    case 'rejectSignup': {
      need(s, ['Admin']); ssheet(); const x = rows('Signups').find(y => y.ID === r.id); if (!x || x.Status !== 'Pending') throw new Error('This request is no longer pending');
      const sg = SS.getSheetByName('Signups'); sg.getRange(x._row, 7).setValue(''); sg.getRange(x._row, 10).setValue('Rejected'); sg.getRange(x._row, 13).setValue(String(r.reason || '').slice(0, 200));
      const emailed = mailTo(x.Email, 'Your account request', '<div style="font-family:Arial,sans-serif;max-width:560px;color:#10243a"><p>Hello ' + EH(x.Name) + ', your account request was not approved.' + (r.reason ? '<br>Reason: ' + EH(r.reason) : '') + '</p><p>Please contact the administrator if you need help.</p></div>');
      return {emailed};
    }
    case 'applyExisting': {
      need(s, ['Admin']); ssheet(); const x = rows('Signups').find(y => y.ID === r.id); if (!x || x.Status !== 'Pending' || x.Kind !== 'existing') throw new Error('This request is no longer pending');
      const u = rows('Users').find(y => String(y.Username) === String(x.ExistingUser)); if (!u) throw new Error('The existing account was not found');
      const us = SS.getSheetByName('Users'); us.getRange(u._row, 2).setValue(x.Name); us.getRange(u._row, 8, 1, 2).setValues([[x.Email, x.Phone]]);
      const sg = SS.getSheetByName('Signups'); sg.getRange(x._row, 10).setValue('Approved'); sg.getRange(x._row, 13).setValue('Existing account details updated');
      const emailed = mailTo(x.Email, 'Your account details were updated', '<div style="font-family:Arial,sans-serif;max-width:560px;color:#10243a"><p>Hello ' + EH(x.Name) + ', the administrator updated the contact details on your account <b>' + EH(u.Username) + '</b>. Your password was not changed.</p></div>');
      return {emailed};
    }
    case 'dismissSignup': { need(s, ['Admin']); ssheet(); const x = rows('Signups').find(y => y.ID === r.id); if (!x) throw new Error('Request not found'); const sg = SS.getSheetByName('Signups'); sg.getRange(x._row, 7).setValue(''); sg.getRange(x._row, 10).setValue('Dismissed'); return true; }
    case 'smsRaw': need(s, ['Admin']); return gatewaySms(r.to, r.message);
    case 'changePassword': {
      const u = rows('Users').find(x => String(x.Username) === s.username);
      if (!u || String(u.Hash).trim() !== hash(String(r.old || ''))) throw new Error('The current password is wrong');
      const pw = String(r.pw || ''); if (pw.length < 8) throw new Error('New password must be at least 8 characters'); if (pw === String(r.old)) throw new Error('Choose a different password');
      const nh = hash(pw), us = SS.getSheetByName('Users'); us.getRange(u._row, 4).setValue(nh); us.getRange(u._row, 10).setValue('');
      return {token: makeToken({Username: u.Username, Hash: nh})};
    }
    case 'resetPassword': {
      need(s, ['Admin']); const u = rows('Users').find(x => String(x.Username).toLowerCase() === String(r.username || '').toLowerCase()); if (!u) throw new Error('User not found');
      const ab = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789', dg = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Date.now()); let temp = '';
      for (let i = 0; i < 10; i++) temp += ab.charAt((dg[i] & 255) % ab.length);
      const us = SS.getSheetByName('Users'); us.getRange(u._row, 4).setValue(hash(temp)); us.getRange(u._row, 10).setValue('Y');
      if (r.id) { ssheet(); const x = rows('Signups').find(y => y.ID === r.id); if (x) { SS.getSheetByName('Signups').getRange(x._row, 10).setValue('Approved'); SS.getSheetByName('Signups').getRange(x._row, 13).setValue('Temporary password sent'); } }
      const cf = config(), link = /^https:\/\//.test(r.link || '') ? r.link : '';
      const emailed = mailTo(u.Email, 'Your temporary password - ' + (cf.Company || 'IS PROJECTS'), '<div style="font-family:Arial,sans-serif;max-width:560px;color:#10243a"><h3>' + EH(cf.Company || 'IS PROJECTS') + '</h3><p>Hello ' + EH(u.Name) + ', the administrator reset your password.</p><p>Username: <b>' + EH(u.Username) + '</b><br>Temporary password: <b>' + EH(temp) + '</b></p><p>You will be asked to choose a new password when you sign in.</p>' + (link ? '<p><a href="' + EH(link) + '">Open the system</a></p>' : '') + '</div>');
      return {temp, emailed, username: u.Username, name: u.Name, email: u.Email || '', phone: u.Phone || ''};
    }
    case 'export': need(s, ['Admin']); return exportAll();
    case 'transferSql': need(s, ['Admin']); return transferSql(r.data || {});
    case 'prfList': {
      psheet(); const out = [];
      rows('PRF').forEach(x => { const d = pdef(pdata(x)); if (pvis(s, x, d)) out.push({ID: x.ID, Date: x.Date, Area: x.Area, Branch: x.Branch, Reason: x.Reason, Status: x.Status, CreatedBy: x.CreatedBy, Data: d}); });
      const pc = prfCfg(), cf = config();
      return {rows: out, groups: pc.groups, flow: pc.flow, branches: cf.Branch, company: cf.Company, users: adm ? rows('Users').map(u => ({username: u.Username, name: u.Name, role: u.Role, dept: u.Dept})) : []};
    }
    case 'prfCreate': {
      need(s, ['Admin','ISD','User','Approver','Reviewer']); const d0 = r.data || {}, br = d0.branch;
      if (!inScope(s, {Branch: br, Area: areaOf(br)})) throw new Error('You can only create requests for your assigned branches');
      if (!/^\d{4}-\d{2}-\d{2}$/.test(String(d0.date || ''))) throw new Error('Enter the date of request');
      const reason = String(d0.reason || '').trim(); if (!reason) throw new Error('Enter the reason for the request');
      const items = cleanItems(d0.items), psh = psheet(), pre = 'PRF-' + new Date().getFullYear() + '-'; let mx = 0;
      rows('PRF').forEach(x => { if (String(x.ID).indexOf(pre) === 0) mx = Math.max(mx, parseInt(String(x.ID).slice(pre.length)) || 0); });
      const d = pdef({items}); d.sign.requested = sig(s);
      const rec = {ID: pre + ('0000' + (mx + 1)).slice(-4), Date: d0.date, Area: areaOf(br), Branch: br, Reason: reason.slice(0, 1000), CreatedBy: s.username, Data: JSON.stringify(d)}; rec.Status = pstat(rec, d);
      psh.appendRow(PH.map(k => rec[k] === undefined ? '' : rec[k])); prfNotify(rec, 'Approver', 'New procurement request waiting for branch manager approval'); return rec.ID;
    }
    case 'prfEdit': {
      need(s, ['Admin','ISD','User','Approver','Reviewer']); const rec = findS(s, r.id), d = pdef(pdata(rec)), d0 = r.data || {};
      if (s.role === 'User' && !(rec.CreatedBy === s.username && rec.Status === 'Pending Branch Approval')) throw new Error('You can only edit your own request while it waits for branch approval');
      if (d.flow.forwarded || d.canvass.length || (d.sign.assessor && !adm)) throw new Error('This request is locked for editing');
      if (!inScope(s, {Branch: d0.branch, Area: areaOf(d0.branch)})) throw new Error('Branch is outside your assigned branches');
      if (!/^\d{4}-\d{2}-\d{2}$/.test(String(d0.date || ''))) throw new Error('Enter the date of request');
      if (!String(d0.reason || '').trim()) throw new Error('Enter the reason for the request');
      d.items = cleanItems(d0.items); d.diag.items = {}; rec.Date = d0.date; rec.Branch = d0.branch; rec.Area = areaOf(d0.branch); rec.Reason = String(d0.reason).trim().slice(0, 1000); d.sign.assessor = null; psave(rec, d); return true;
    }
    case 'prfDelete': {
      need(s, ['Admin','ISD','User']); const rec = findS(s, r.id), own = rec.CreatedBy === s.username;
      if (!((adm || own) && rec.Status === 'Pending Branch Approval') && !((adm || s.role === 'ISD') && ['For ISD Assessment','For Forwarding'].indexOf(rec.Status) >= 0)) throw new Error('This request can no longer be deleted');
      purgeFiles(f => f.RecordID === rec.ID); psheet().deleteRow(rec._row); return true;
    }
    case 'prfItemsEdit': {
      need(s, ['Admin','ISD']);
      const rec = findS(s, r.id), d = pdef(pdata(rec)), d0 = r.data || {};
      if (d.rejected || d.sign.purchaser || d.sign.checker) throw new Error('Requested items are locked after purchase or completion');
      if (Object.keys(d.award || {}).length) throw new Error('Requested items are locked after an award');
      const newItems = cleanItems(d0.items);
      if (newItems.length < d.items.length) throw new Error('Existing requested items cannot be removed here. You may add more items.');
      d.items = newItems;
      d.diag = {items: {}, remarks: ''};
      d.sign.assessor = null;
      psave(rec, d);
      return true;
    }
    case 'prfSaveCfg': {
      need(s, ['Admin']); const g = r.data.groups, f = r.data.flow;
      if (!g || typeof g !== 'object' || !Object.keys(g).length) throw new Error('Add at least one item group');
      if (!f || !Array.isArray(f.steps)) throw new Error('Invalid approval flow');
      cfgSet('PRFGroups', JSON.stringify(g)); cfgSet('PRFFlow', JSON.stringify({name: String(f.name || 'PRF').trim().slice(0, 40) || 'PRF', steps: f.steps.map(x => ({k: String(x.k).slice(0, 12), l: String(x.l).slice(0, 60), u: (x.u || []).map(String)})), acct: String(f.acct || 'Accounting'), diag: (f.diag || []).map(String), award: (f.award || []).map(String)})); return true;
    }
    case 'prfAct': {
      const rec = findS(s, r.id), d = pdef(pdata(rec)), v = r.data || {}, fl = prfCfg().flow, nI = d.items.length;
      // Credentialed PRF actions deliberately authenticate the person signing this specific step.
      // The credentials are verified server-side and are never stored in the PRF.
      const credentialed = ['diagApprove','awardSign','awardApprove'].indexOf(String(r.type)) >= 0;
      const actor = credentialed ? verifyCredential(v.username, v.password) : s;
      const me = String(actor.username).toLowerCase();
      const adm = String(actor.role || '').toLowerCase() === 'admin', deny = () => { throw new Error('Not allowed for your role'); }, bad = m => { throw new Error(m); };
      const isd = adm || actor.role === 'ISD', acct = adm || (String(actor.role || '').toLowerCase() === 'dept' && String(actor.dept || '').trim().toLowerCase() === String(fl.acct || 'Accounting').trim().toLowerCase());
      const inUsers = list => adm || (Array.isArray(list) && list.some(u => String(u).toLowerCase() === me));
      const diagOk = adm || (Array.isArray(fl.diag) && fl.diag.length ? inUsers(fl.diag) : actor.role === 'ISD');
      const awardOk = adm || (Array.isArray(fl.award) && fl.award.length ? inUsers(fl.award) : actor.role === 'ISD');
      const stepOk = k => adm || ((fl.steps.find(x => x.k === k) || {u: []}).u || []).some(u => String(u).toLowerCase() === me);
      if (d.rejected || d.sign.checker) bad('This request is closed'); let ret = null;
      switch (r.type) {
        case 'branch': if (!(adm || s.role === 'Approver')) deny(); if (d.sign.branch) bad('Already approved'); d.sign.branch = sig(s); if (v.remarks) d.rem.branch = String(v.remarks).slice(0, 500); break;
        case 'diag': {
          if (!diagOk) deny(); if (!d.sign.branch) bad('Waiting for branch manager approval'); const it = {};
          Object.keys(v.items || {}).forEach(i => { const x = v.items[i]; if (+i >= 0 && +i < nI && x && x.group) it[i] = {group: String(x.group).slice(0, 60), parts: x.parts || {}, remarks: String(x.remarks || '').slice(0, 500)}; });
          d.diag = {items: it, remarks: String(v.remarks || '').slice(0, 1000)}; d.sign.assessor = sig(s); delete d.sign.diagApproved; break;
        }
        case 'assess': if (!diagOk) deny(); if (!d.sign.branch) bad('Waiting for branch manager approval'); if (!Object.keys(d.diag.items).length) bad('Complete the diagnostic first'); d.sign.assessor = sig(actor); delete d.sign.diagApproved; break;
        case 'diagApprove': {
          if (!(adm || actor.role === 'Approver' || actor.role === 'ISD')) deny();
          if (!d.sign.branch) bad('Branch Manager must approve the PRF before diagnostic approval');
          if (!d.sign.assessor) bad('Diagnostic must first be prepared / assessed by ISD');
          if (d.sign.diagApproved && !adm) bad('Diagnostic is already approved');
          d.sign.diagApproved = sig(actor); break;
        }
        case 'forward': if (!isd) deny(); if (!d.sign.branch) bad('Waiting for branch manager approval'); if (!d.sign.assessor) bad('Prepared / Assessed by must be recorded before forwarding'); if (!d.sign.diagApproved) bad('Diagnostic approval is required before forwarding'); if (d.flow.forwarded) bad('Already forwarded');
          d.flow.steps = fl.steps.map(x => x.k).filter(k => (v.steps || []).indexOf(k) >= 0); d.flow.forwarded = sig(s); break;
        case 'charges': if (!acct) deny(); if (d.sign.verifier && !adm) bad('Already verified');
          (v.charges || []).forEach((c, i) => { if (d.items[i]) d.items[i].c = (c === '' || c == null) ? null : Number(c); }); d.rem.acct = String(v.remarks || '').slice(0, 1000); break;
        case 'verify': if (!acct) deny(); d.sign.verifier = sig(s); break;
        case 'approve': {
          const k = v.key, idx = d.flow.steps.indexOf(k);
          if (idx < 0) bad('Not part of this request');
          if (!d.sign.verifier) bad('Accounting verification is required before approval');
          if (!stepOk(k)) bad('You are not assigned to this approval');
          if (d.sign.approvals[k]) bad('Already signed');
          for (let i = 0; i < idx; i++) if (!d.sign.approvals[d.flow.steps[i]]) bad('Waiting for ' + d.flow.steps[i] + ' first');
          d.sign.approvals[k] = sig(s); if (v.remarks) d.rem['ap_' + k] = String(v.remarks).slice(0, 500); break;
        }
        case 'reject': if (!(isd || acct || (s.role === 'Approver' && !d.sign.branch) || d.flow.steps.some(stepOk))) deny(); if (d.sign.purchaser) bad('Already purchased'); d.rejected = {by: sig(s), reason: String(v.reason || '').slice(0, 300)}; break;
        case 'canvass': {
          if (!(adm || rec.CreatedBy.toLowerCase() === me || isd)) deny(); const i = Number(v.item); if (!(i >= 0 && i < nI)) bad('Choose the requested item');
          const store = String(v.store || '').trim(), price = Number(v.price), nm = String(v.name || '').trim();
          if (!store || !nm) bad('Item and store are required'); if (!(price >= 0)) bad('Enter the price'); if (!/^\d{4}-\d{2}-\d{2}$/.test(String(v.date || ''))) bad('Enter the canvass date');
          let e = v.id ? d.canvass.find(x => x.id === v.id) : null;
          if (e && d.award[e.item] === e.id) bad('This canvass is awarded and cannot be modified.');
          if (!e) { if (d.canvass.length >= 60) bad('Too many canvass entries'); e = {id: 'C' + Date.now().toString(36)}; d.canvass.push(e); }
          Object.assign(e, {item: i, name: nm.slice(0, 160), date: v.date, store: store.slice(0, 120), price, desc: String(v.desc || '').slice(0, 300), remarks: String(v.remarks || '').slice(0, 300), by: sig(s)});
          d.sign.canvasser = sig(s); ret = e.id; break;
        }
        case 'canvassDel': {
          if (!(adm || rec.CreatedBy.toLowerCase() === me || isd)) deny();
          const e = d.canvass.find(x => x.id === v.id); if (!e) bad('Canvass entry not found');
          const awarded = d.award[e.item] === e.id;
          if (awarded && !adm) bad('This canvass has been awarded and cannot be deleted.');
          if (d.sign.purchaser && !adm) bad('Already purchased');
          if (awarded && adm && String(v.force || '') !== '1') bad('This canvass is awarded. Confirm Force Delete as Admin.');
          if (awarded) { delete d.award[e.item]; delete d.sign.awarded; }
          d.canvass = d.canvass.filter(x => x.id !== v.id);
          purgeFiles(f => f.RecordID === rec.ID && f.Ref === 'cv:' + v.id);
          break;
        }
        case 'award': {
          if (!awardOk) deny();
          if (d.sign.purchaser) bad('Already purchased');
          const i = Number(v.item); if (!(i >= 0 && i < nI)) bad('Choose the requested item');
          if (!v.id) delete d.award[i]; else { const e = d.canvass.find(x => x.id === v.id && x.item === i); if (!e) bad('Choose a canvass entry for this item'); d.award[i] = e.id; }
          if (!d.items.every((x, k) => d.award[k])) { delete d.sign.awarded; delete d.sign.awardApproved; } break;
        }
        case 'awardSign': {
          if (!awardOk) deny();
          if (!d.items.every((x, k) => d.award[k])) bad('Award every requested item first');
          if (d.sign.awarded && !adm) bad('Awarded by is already recorded');
          d.sign.awarded = sig(actor); d.audit.awarded = sig(actor); delete d.sign.awardApproved; break;
        }
        case 'awardApprove': {
          const ism = fl.steps.find(x => String(x.k || '').toUpperCase() === 'ISM');
          const ismOk = ism && Array.isArray(ism.u) && ism.u.some(u => String(u).toLowerCase() === me);
          if (!(adm || actor.role === 'Approver' || actor.role === 'ISD' || ismOk)) deny();
          if (!d.sign.awarded) bad('Awarded by must be recorded first');
          if (d.sign.awardApproved && !adm) bad('Award is already approved');
          d.sign.awardApproved = sig(actor); break;
        }
        case 'buy': if (d.sign.checker) bad('This request has already been checked and is locked'); if (d.sign.purchaser && !adm) bad('Already signed'); Object.keys(v.rows || {}).forEach(i => { if (d.award[i]) d.buy[i] = {or: String(v.rows[i].or || '').slice(0, 60), store: String(v.rows[i].store || '').slice(0, 120)}; }); d.audit.purchase = sig(s); break;
        case 'purchaser': if (d.sign.checker) bad('This request has already been checked and is locked'); if (!d.items.every((x, k) => d.award[k])) bad('Award every item first'); if (!d.sign.awardApproved) bad('Award approval is required before purchase'); if (!d.items.every((x, k) => d.buy[k] && d.buy[k].or && d.buy[k].store)) bad('Enter the OR number and store for every item'); d.sign.purchaser = sig(s); d.audit.purchaser = sig(s); break;
        case 'dates': {
          if (d.sign.checker) bad('This request has already been checked and is locked');
          const compiled = v.compiled, received = v.received, transmittal = v.transmittal;
          if (compiled !== undefined && compiled && !/^\d{4}-\d{2}-\d{2}$/.test(compiled)) bad('Invalid compiled date');
          if (received !== undefined && received && !/^\d{4}-\d{2}-\d{2}$/.test(received)) bad('Invalid received date');
          if (transmittal !== undefined && transmittal && !/^\d{4}-\d{2}-\d{2}$/.test(transmittal)) bad('Invalid transmittal date');
          if (received !== undefined && received && !d.sign.purchaser) bad('Record the Purchaser first before entering the date received');
          if (compiled !== undefined && compiled && !d.dates.received && !(received)) bad('Enter the date received before entering the date compiled');
          if (transmittal !== undefined && transmittal && !(compiled || d.dates.compiled)) bad('Enter the date compiled before entering the transmittal date');
          if (received !== undefined && received && compiled && compiled < received) bad('Date compiled cannot be earlier than date received');
          if (transmittal !== undefined && transmittal && (compiled || d.dates.compiled) > transmittal) bad('Transmittal date cannot be earlier than date compiled');
          [['received', received], ['compiled', compiled], ['transmittal', transmittal]].forEach(([k, val]) => {
            if (val === undefined) return;
            d.dates[k] = val || '';
            if (val) d.audit[k] = sig(s); else delete d.audit[k];
          });
          break;
        }
        case 'isnotation': {
          if (!isd) deny();
          const txt = String(v.remarks || '').trim();
          if (!txt) bad('Enter the IS notation.');
          d.rem.isNotation = txt.slice(0, 2000);
          d.sign.isNotation = sig(s);
          d.audit.isNotation = sig(s);
          break;
        }
        case 'notation': if (!isd) deny(); if (!d.sign.purchaser || !d.dates.received) bad('Enter the date received first'); d.sign.notation = sig(s); d.rem.notation = String(v.remarks || '').slice(0, 1000); break;
        case 'check': if (!(adm || s.role === 'Reviewer')) deny(); if (!d.sign.notation) bad('Waiting for the ISD notation'); d.sign.checker = sig(s); d.rem.checker = String(v.remarks || '').slice(0, 1000); break;
        default: bad('Unknown step');
      }
      psave(rec, d);
      if (r.type === 'branch') prfNotify(rec, 'ISD', 'Approved by the branch manager - ready for ISD assessment');
      if (r.type === 'assess') prfNotify(rec, 'ISD', 'PRF assessment completed - ready for forwarding');
      if (r.type === 'forward') prfNotify(rec, 'Dept', 'A procurement request has been forwarded for your action');
      return {status: rec.Status, id: ret};
    }
    case 'saveLogin': {
      need(s, ['Admin']); const d = r.data;
      if (d.img !== undefined && d.img !== null) setImg('login', d.img);
      cfgSet('Login', JSON.stringify(d.login || {})); return true;
    }
    case 'saveHome': {
      need(s, ['Admin']); const sheet = isheet(), d = r.data;
      Object.keys(d.images || {}).forEach(k => {
        rows('Images').filter(x => x.Key === k).sort((a, b) => b._row - a._row).forEach(x => sheet.deleteRow(x._row));
        const v = String(d.images[k] || ''); if (!v) return;
        if (v.indexOf('data:image/') !== 0 || v.length > 240000) throw new Error('Image is too large. Use a smaller picture.');
        for (let i = 0; i * 45000 < v.length; i++) sheet.appendRow([k, i, v.slice(i * 45000, (i + 1) * 45000)]);
      });
      cfgSet('Home', JSON.stringify(d.home || {})); return true;
    }
    case 'saveConfig': {
      need(s, ['Admin']); const c = SS.getSheetByName('Config'), d = r.data, v = [CH];
      d.Branch.forEach(b => v.push(['Branch', b.name, b.area])); d.ErrorType.forEach(e => v.push(['ErrorType', e.name, e.mon || ''])); d.Dept.forEach(x => v.push(['Dept', x, ''])); (d.Monitor || []).forEach(m => v.push(['Monitor', m.code, m.name]));
      Object.keys(d.Fields || {}).forEach(k => { if (d.Fields[k].length) v.push(['Fields', k, JSON.stringify(d.Fields[k])]); });
      rows('Config').filter(x => x.Type === 'Setting').forEach(x => v.push(['Setting', x.Value, x.Area]));
      c.clear(); c.getRange(1, 1, v.length, 3).setValues(v); return true;
    }
    case 'listUsers': need(s, ['Admin']); return rows('Users').map(u => ({username: u.Username, name: u.Name, role: u.Role, active: String(u.Active).toUpperCase() !== 'FALSE', branch: u.Branch, dept: u.Dept, email: u.Email || '', phone: u.Phone || ''}));
    case 'saveUser': {
      need(s, ['Admin']); const us = SS.getSheetByName('Users'), d = r.data;
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(d.email || '').trim())) throw new Error('A valid email address is required');
      d.phone = String(d.phone || '').replace(/[^\d+]/g, ''); if (!/^\+?\d{10,13}$/.test(d.phone)) throw new Error('A valid contact number is required');
      const ex = rows('Users').find(x => String(x.Username).toLowerCase() === d.username.toLowerCase());
      const row = [d.username, d.name, d.role, d.password ? hash(d.password) : (ex ? ex.Hash : hash('changeme')), d.active !== false, d.branch || '', d.dept || '', d.email.trim(), d.phone, ex ? (ex.MustChange || '') : ''];
      if (ex) us.getRange(ex._row, 1, 1, 10).setValues([row]); else us.appendRow(row); return true;
    }
  }
  throw new Error('Unknown action');
}
