/* IS PROJECTS - ISRF backend v2. After pasting: run repair() once, then Deploy > Manage deployments > New version. Login: admin / admin123 */
const SS = SpreadsheetApp.getActiveSpreadsheet();
const H = ['ID','Date','Area','Branch','ErrorType','DocsDamage','ProblemDetails','RequestedBy','ApprovedBy','Decision','ForwardTo','Notes','AnsweredBy','ReviewedBy','Status','CreatedBy','Monitoring','Data'];
const UH = ['Username','Name','Role','Hash','Active','Branch','Dept','Email'], CH = ['Type','Value','Area'];

function sh(n, h) { if (!n || !h) throw new Error('Do not run sh() directly.'); let s = SS.getSheetByName(n); if (!s) { s = SS.insertSheet(n); s.appendRow(h); s.setFrozenRows(1); } return s; }
function hash(p) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'isrf:' + p).map(b => ('0' + (b & 255).toString(16)).slice(-2)).join(''); }
function setup() {
  sh('ISRF', H).getRange('B:B').setNumberFormat('@');
  const u = sh('Users', UH); u.getRange('A:D').setNumberFormat('@');
  if (u.getLastRow() < 2) u.appendRow(['admin','Administrator','Admin',hash('admin123'),true,'','','']);
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
  const row = ['admin','Administrator','Admin',hash('admin123'),true,'','',''];
  if (ex) u.getRange(ex._row, 1, 1, 8).setValues([row]); else u.appendRow(row);
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
    if (name === 'ISRF' && hdr[i] === 'Monitoring' && s.getLastRow() > 1) s.getRange(2, i + 1, s.getLastRow() - 1, 1).setValue('ISRF');
  }
}
function upgrade() { upg('ISRF', H); upg('Users', UH); }
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
  c.Settings = settings(); c.SmsOn = !!PropertiesService.getScriptProperties().getProperty('SMS_KEY');
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
function save(rec) { SS.getSheetByName('ISRF').getRange(rec._row, 1, 1, H.length).setValues([H.map(k => rec[k] === undefined ? '' : rec[k])]); }
function find(id) { const r = rows('ISRF').find(x => x.ID === id); if (!r) throw new Error('Record not found'); return r; }
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
function findS(s, id) { const r = find(id); if (!canSee(s, r)) throw new Error('This request is outside your assigned branches'); return r; }

/* ---- attachments (stored in one Google Drive folder, metadata in the Files tab) ---- */
const SET = {MaxFileMB: 2, MaxFiles: 5, CapMB: 3000, RetainDays: 365}, FH = ['FileId','RecordID','Name','Size','Mime','UploadedBy','UploadedAt'];
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
  return {username: u.Username, name: u.Name, role: u.Role, branch: u.Branch || '', dept: u.Dept || ''};
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

function handle(r) {
  const cache = CacheService.getScriptCache();
  if (r.action === 'login') {
    const u = rows('Users').find(x => String(x.Username).toLowerCase() === String(r.username).toLowerCase());
    if (!u || String(u.Hash).trim() !== hash(String(r.password)) || String(u.Active).toUpperCase() === 'FALSE') throw new Error('Wrong username or password');
    const s = {username: u.Username, name: u.Name, role: u.Role, branch: u.Branch || '', dept: u.Dept || ''}; return {token: makeToken(u), user: s};
  }
  if (r.action === 'brand') {
    const b = config(); let L = {}; try { L = JSON.parse(cfgGet('Login') || '{}'); } catch (e) {}
    return {Company: b.Company, Logo: b.Logo, UI: b.UI, Login: L, LoginImg: imgGet('login')};
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
    case 'files': { findS(s, r.id); return frows().filter(f => f.RecordID === r.id).map(f => ({FileId: f.FileId, Name: f.Name, Size: f.Size, UploadedBy: f.UploadedBy, UploadedAt: f.UploadedAt})); }
    case 'upload': {
      need(s, ['Admin','ISD','User','Approver','Dept','Reviewer']);
      const rec = findS(s, r.id), st = status(rec); if (!adm && (st === 'Completed' || st === 'Rejected')) throw new Error('This request is closed');
      const set = settings(), ext = String(r.name).split('.').pop().toLowerCase();
      if (['pdf','jpg','jpeg','png'].indexOf(ext) < 0) throw new Error('Only PDF, JPG and PNG files are allowed');
      const bytes = Utilities.base64Decode(r.b64), fr = frows();
      if (bytes.length > set.MaxFileMB * 1048576) throw new Error('File is larger than ' + set.MaxFileMB + ' MB');
      if (fr.filter(f => f.RecordID === rec.ID).length >= set.MaxFiles) throw new Error('Maximum ' + set.MaxFiles + ' files per request');
      if (fr.reduce((a, f) => a + Number(f.Size || 0), 0) + bytes.length > set.CapMB * 1048576) throw new Error('Storage limit reached. Ask the administrator to clean up old attachments.');
      const file = folder().createFile(Utilities.newBlob(bytes, r.mime || 'application/octet-stream', rec.ID + '_' + r.name));
      fsheet().appendRow([file.getId(), rec.ID, r.name, bytes.length, r.mime || '', s.username, new Date().toISOString()]); return true;
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
    case 'sms': {
      need(s, ['Admin','ISD','User','Approver','Dept','Reviewer']); findS(s, r.id);
      const P = PropertiesService.getScriptProperties(), key = P.getProperty('SMS_KEY'); if (!key) throw new Error('SMS gateway is not configured');
      const nums = String(r.to || '').split(/[,;\s]+/).filter(x => /^\+?\d{10,13}$/.test(x)); if (!nums.length || nums.length > 5) throw new Error('Enter 1 to 5 valid mobile numbers');
      const pl = {apikey: key, number: nums.join(','), message: String(r.message || '').slice(0, 459)}; if (P.getProperty('SMS_SENDER')) pl.sendername = P.getProperty('SMS_SENDER');
      const res = UrlFetchApp.fetch('https://api.semaphore.co/api/v4/messages', {method: 'post', payload: pl, muteHttpExceptions: true});
      if (res.getResponseCode() >= 300) throw new Error('SMS gateway error: ' + res.getContentText().slice(0, 120)); return nums.length;
    }
    case 'saveSms': {
      need(s, ['Admin']); const P = PropertiesService.getScriptProperties();
      if (r.clear) { P.deleteProperty('SMS_KEY'); P.deleteProperty('SMS_SENDER'); return true; }
      if (r.key) P.setProperty('SMS_KEY', String(r.key).trim()); P.setProperty('SMS_SENDER', String(r.sender || '').trim()); return true;
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
    case 'listUsers': need(s, ['Admin']); return rows('Users').map(u => ({username: u.Username, name: u.Name, role: u.Role, active: String(u.Active).toUpperCase() !== 'FALSE', branch: u.Branch, dept: u.Dept, email: u.Email || ''}));
    case 'saveUser': {
      need(s, ['Admin']); const us = SS.getSheetByName('Users'), d = r.data, ex = rows('Users').find(x => String(x.Username).toLowerCase() === d.username.toLowerCase());
      const row = [d.username, d.name, d.role, d.password ? hash(d.password) : (ex ? ex.Hash : hash('changeme')), d.active !== false, d.branch || '', d.dept || '', d.email || ''];
      if (ex) us.getRange(ex._row, 1, 1, 8).setValues([row]); else us.appendRow(row); return true;
    }
  }
  throw new Error('Unknown action');
}
