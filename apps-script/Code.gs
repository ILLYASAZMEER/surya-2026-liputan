/**
 * SURYA 2026 — API Jadual Liputan (Pasukan Coverage, Media Sosial JKR Sarawak)
 * Google Apps Script web app. Data disimpan dalam Google Sheet "SURYA 2026 – Data Website Liputan".
 *
 * Langkah sekali sahaja:
 *   1. Jalankan fungsi setup() (pilih "setup" di bar atas, tekan Run, benarkan akses).
 *   2. Deploy > New deployment > Web app: Execute as "Me", Who has access "Anyone".
 * Tukar PIN admin: Project Settings > Script Properties > ADMIN_PIN.
 */

const TZ = 'Asia/Kuching';
const SLOT_RE = /^[\w-]{1,80}$/;
const ID_RE = /^[\w-]{1,40}$/;
const HEADERS = {
  Petugas: ['id', 'nama', 'telefon', 'peranan', 'urutan', 'foto', 'avail', 'unit', 'kamera'],
  Tugasan: ['slot_id', 'slot', 'petugas_id', 'petugas_nama'],
  Status: ['slot_id', 'slot', 'drive', 'drive_oleh', 'drive_masa', 'publisiti', 'publisiti_oleh', 'publisiti_masa', 'nota', 'nota_oleh', 'nota_masa', 'tiada_publisiti', 'tiada_publisiti_oleh', 'tiada_publisiti_masa', 'priority', 'priority_oleh', 'priority_masa']
};

function setup() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ADMIN_PIN')) props.setProperty('ADMIN_PIN', '(ditetapkan dalam Script Properties)');
  const ss = ss_();
  Logger.log('Google Sheet data: ' + ss.getUrl());
}

/* ---------------- sheet helpers ---------------- */
function ss_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* recreate below */ }
  }
  const ss = SpreadsheetApp.create('SURYA 2026 – Data Website Liputan');
  Object.keys(HEADERS).forEach(function (name) {
    const sh = ss.insertSheet(name);
    sh.getRange('A:Z').setNumberFormat('@');
    sh.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  ss.getSheets().forEach(function (sh) { if (!HEADERS[sh.getName()]) ss.deleteSheet(sh); });
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}
function sheet_(name) { return ss_().getSheetByName(name); }
function rows_(name) {
  const sh = sheet_(name);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const h = HEADERS[name];
  return sh.getRange(2, 1, last - 1, h.length).getValues()
    .filter(function (r) { return r[0] !== ''; })
    .map(function (r) { const o = {}; h.forEach(function (k, i) { o[k] = String(r[i]); }); return o; });
}
function findRow_(name, key) {
  const sh = sheet_(name);
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === key) return i + 2;
  return -1;
}
function upsert_(name, key, obj) {
  const sh = sheet_(name), h = HEADERS[name];
  let row = findRow_(name, key);
  let cur = {};
  if (row > 0) {
    const v = sh.getRange(row, 1, 1, h.length).getValues()[0];
    h.forEach(function (k, i) { cur[k] = String(v[i]); });
  } else {
    row = Math.max(sh.getLastRow(), 1) + 1;
  }
  Object.keys(obj).forEach(function (k) { cur[k] = obj[k]; });
  cur[h[0]] = key;
  sh.getRange(row, 1, 1, h.length).setNumberFormat('@').setValues([h.map(function (k) { return cur[k] == null ? '' : String(cur[k]); })]);
}
function now_() { return Utilities.formatDate(new Date(), TZ, 'd/M/yyyy HH:mm'); }
function bump_(key) { PropertiesService.getScriptProperties().setProperty(key, String(Date.now())); }
function rev_(key) { return PropertiesService.getScriptProperties().getProperty(key) || '0'; }
function str_(v, max) { return String(v == null ? '' : v).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '').slice(0, max); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ---------------- read ---------------- */
function petugasData_() {
  return rows_('Petugas').map(function (r) {
    return { id: r.id, nama: r.nama, telefon: r.telefon, peranan: r.peranan, urutan: Number(r.urutan) || 0, foto: /^data:image\/(jpeg|png|webp);base64,/.test(r.foto) ? r.foto : '', avail: r.avail || '', unit: r.unit || '', kamera: r.kamera || '' };
  });
}
function slotData_() {
  const tug = {}, st = {};
  rows_('Tugasan').forEach(function (r) { tug[r.slot_id] = r.petugas_id ? r.petugas_id.split(',').filter(String) : []; });
  rows_('Status').forEach(function (r) {
    st[r.slot_id] = {
      drive: r.drive ? { by: r.drive_oleh, at: r.drive_masa } : null,
      publisiti: r.publisiti ? { by: r.publisiti_oleh, at: r.publisiti_masa } : null,
      nota: r.nota ? { t: r.nota, by: r.nota_oleh, at: r.nota_masa } : null,
      tiada_publisiti: r.tiada_publisiti ? { by: r.tiada_publisiti_oleh, at: r.tiada_publisiti_masa } : null,
      priority: r.priority ? { by: r.priority_oleh, at: r.priority_masa } : null
    };
  });
  return { tugasan: tug, status: st };
}
function readAll_(haveRev, havePrev) {
  const rev = rev_('REV'), prev = rev_('PREV');
  const res = { ok: true, rev: rev, prev: prev };
  const cache = CacheService.getScriptCache();
  if (haveRev !== rev) {
    let d = cache.get('d' + rev);
    if (d) d = JSON.parse(d); else { d = slotData_(); try { cache.put('d' + rev, JSON.stringify(d), 600); } catch (e) {} }
    res.tugasan = d.tugasan; res.status = d.status;
  }
  if (havePrev !== prev) res.petugas = petugasData_();
  return res;
}

/* ---------------- write ---------------- */
function checkPin_(pin) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('pinfail') || 0);
  if (fails >= 10) throw new Error('locked');
  const ok = String(pin || '') !== '' && String(pin) === PropertiesService.getScriptProperties().getProperty('ADMIN_PIN');
  if (!ok) cache.put('pinfail', String(fails + 1), 600);
  return ok;
}
function requirePin_(pin) { if (!checkPin_(pin)) throw new Error('pin'); }

function tick_(b) {
  if (!SLOT_RE.test(b.slot) || ['drive', 'publisiti', 'tiada_publisiti', 'priority'].indexOf(b.key) < 0) throw new Error('bad');
  ensureHeader_('Status');
  const o = { slot: str_(b.label, 200) };
  o[b.key] = b.on ? '✓' : '';
  o[b.key + '_oleh'] = b.on ? str_(b.by, 60) : '';
  o[b.key + '_masa'] = b.on ? now_() : '';
  upsert_('Status', b.slot, o);
  bump_('REV');
  return { ok: true };
}
function nota_(b) {
  if (!SLOT_RE.test(b.slot)) throw new Error('bad');
  const t = str_(b.t, 2000);
  upsert_('Status', b.slot, { slot: str_(b.label, 200), nota: t, nota_oleh: t ? str_(b.by, 60) : '', nota_masa: t ? now_() : '' });
  bump_('REV');
  return { ok: true };
}
function ensureHeader_(name) {
  const sh = sheet_(name), h = HEADERS[name];
  const cur = sh.getRange(1, 1, 1, h.length).getValues()[0].map(String);
  if (cur.join('|') !== h.join('|')) sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
}
function ensurePetugasHeader_() {
  const sh = sheet_('Petugas'), h = HEADERS.Petugas;
  const cur = sh.getRange(1, 1, 1, h.length).getValues()[0].map(String);
  if (cur.join('|') !== h.join('|')) sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
}
function putPetugas_(p) {
  if (!p || !ID_RE.test(p.id)) throw new Error('bad');
  const foto = /^data:image\/(jpeg|png|webp);base64,[A-Za-z0-9+/=]+$/.test(p.foto || '') && p.foto.length <= 48000 ? p.foto : '';
  const keep = !foto && p.keepFoto ? (rows_('Petugas').filter(function (r) { return r.id === p.id; })[0] || {}).foto || '' : '';
  const avail = /^([1-5]=(ada|pagi|petang|tiada)?;?){0,5}$/.test(p.avail || '') ? String(p.avail || '') : '';
  upsert_('Petugas', p.id, { nama: str_(p.nama, 80), telefon: str_(p.telefon, 20), peranan: str_(p.peranan, 40), urutan: String(Number(p.urutan) || Date.now()), foto: foto || keep, avail: avail, unit: str_(p.unit, 80), kamera: str_(p.kamera, 120) });
}
function savePetugas_(p) {
  ensurePetugasHeader_();
  putPetugas_(p);
  relabel_();
  bump_('PREV'); bump_('REV');
  return { ok: true };
}
function bulkPetugas_(list) {
  if (!Array.isArray(list) || list.length > 80) throw new Error('bad');
  ensurePetugasHeader_();
  list.forEach(putPetugas_);
  relabel_();
  bump_('PREV'); bump_('REV');
  return { ok: true, n: list.length };
}
function delPetugas_(id) {
  if (!ID_RE.test(id)) throw new Error('bad');
  const row = findRow_('Petugas', id);
  if (row > 0) sheet_('Petugas').deleteRow(row);
  rows_('Tugasan').forEach(function (r) {
    const ids = r.petugas_id.split(',').filter(String);
    if (ids.indexOf(id) >= 0) upsert_('Tugasan', r.slot_id, { petugas_id: ids.filter(function (x) { return x !== id; }).join(',') });
  });
  relabel_();
  bump_('PREV'); bump_('REV');
  return { ok: true };
}
function assign_(slots, ids, labels) {
  if (!Array.isArray(slots) || !Array.isArray(ids) || slots.length > 120) throw new Error('bad');
  ids = ids.filter(function (x) { return ID_RE.test(x); });
  slots.forEach(function (s, i) {
    if (!SLOT_RE.test(s)) return;
    upsert_('Tugasan', s, { slot: str_((labels || [])[i], 200), petugas_id: ids.join(',') });
  });
  relabel_();
  bump_('REV');
  return { ok: true };
}
function relabel_() {
  const names = {};
  rows_('Petugas').forEach(function (r) { names[r.id] = r.nama; });
  const sh = sheet_('Tugasan'), last = sh.getLastRow();
  if (last < 2) return;
  const rng = sh.getRange(2, 3, last - 1, 2), v = rng.getValues();
  rng.setValues(v.map(function (r) { return [r[0], String(r[0]).split(',').filter(String).map(function (id) { return names[id] || ''; }).filter(String).join(', ')]; }));
}

/* ---------------- endpoints ---------------- */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.a === 'all') {
    try { return out_(readAll_(String(p.rev || ''), String(p.prev || ''))); }
    catch (err) { return out_({ ok: false, err: String(err.message || err) }); }
  }
  return out_({ ok: true, app: 'SURYA 2026 Jadual Liputan API' });
}
function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (x) { return out_({ ok: false, err: 'bad_json' }); }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return out_({ ok: false, err: 'busy' });
  try {
    switch (b.a) {
      case 'tick': requirePin_(b.pin); return out_(tick_(b));
      case 'nota': requirePin_(b.pin); return out_(nota_(b));
      case 'pin': return out_({ ok: checkPin_(b.pin) });
      case 'savePetugas': requirePin_(b.pin); return out_(savePetugas_(b.p));
      case 'bulkPetugas': requirePin_(b.pin); return out_(bulkPetugas_(b.list));
      case 'delPetugas': requirePin_(b.pin); return out_(delPetugas_(b.id));
      case 'assign': requirePin_(b.pin); return out_(assign_(b.slots, b.ids, b.labels));
    }
    return out_({ ok: false, err: 'unknown' });
  } catch (err) {
    return out_({ ok: false, err: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}
