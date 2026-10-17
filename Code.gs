/**
 * Бекенд дашборду PR-таймлайнів.
 * Вставте цей код у Google Таблицю → Розширення → Apps Script,
 * потім Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone).
 * Після змін коду: Deploy → Manage deployments → ✏️ → Version: New version → Deploy.
 */

const SHEET_NAME = 'cases';
const STAGES = ['eoi','laa','mpnp','verif','nom','pr','aor','bio','pal','med','fd','p1','p2','ecopr'];
const META = ['nick','stream','ee','security','notes'];
const LISTS = ['adr','ghost'];           // багаторазові події, дати через пробіл
const HEADERS = ['id','key','created','updated'].concat(META, STAGES, LISTS);
const STREAMS = ['swm','swo','ies','bis','other'];
const MAX_CREATES_PER_10_MIN = 30;
const MAX_LIST = 20;

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    return sh;
  }
  // додати колонки, яких бракує (міграція старих таблиць)
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const h = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const used = h.filter(Boolean).length;
  const missing = HEADERS.filter(x => h.indexOf(x) < 0);
  if (missing.length) {
    sh.getRange(1, used + 1, 1, missing.length).setValues([missing]).setFontWeight('bold');
  }
  return sh;
}

function headers_(sh) {
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

function fmt_(x) {
  if (x instanceof Date) return Utilities.formatDate(x, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(x == null ? '' : x).trim();
}

function rows_() {
  const sh = sheet_();
  const v = sh.getDataRange().getValues();
  const h = v[0];
  const out = [];
  for (let i = 1; i < v.length; i++) {
    const o = {};
    h.forEach((name, j) => o[name] = v[i][j]);
    if (!o.id) continue;
    o._row = i + 1;
    out.push(o);
  }
  return out;
}

function toCase_(o) {
  const c = { id: String(o.id), created: fmt_(o.created), updated: fmt_(o.updated), s: {} };
  META.forEach(k => c[k] = fmt_(o[k]));
  STAGES.forEach(k => { const d = fmt_(o[k]); if (d) c.s[k] = d; });
  LISTS.forEach(k => c[k] = fmt_(o[k]).split(/[\s,;]+/).filter(Boolean));
  return c;
}

function doGet() {
  return json_({ ok: true, cases: rows_().map(toCase_) });
}

function doPost(e) {
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.website) return json_({ ok: true, id: 'x', key: 'x' }); // пастка для ботів
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      if (b.action === 'create') return json_(create_(b.case));
      if (b.action === 'update') return json_(update_(b.id, b.key, b.case));
      if (b.action === 'delete') return json_(delete_(b.id, b.key));
      return json_({ ok: false, error: 'Невідома дія' });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function date_(d, what) {
  d = String(d || '').trim();
  if (!d) return '';
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) throw new Error('Неправильна дата: ' + what);
  const y = Number(d.slice(0, 4));
  if (y < 2015 || y > 2035) throw new Error('Неправдоподібний рік: ' + d);
  return d;
}

function clean_(c) {
  if (!c || typeof c !== 'object') throw new Error('Порожні дані');
  const s = c.s || {};
  const o = {};
  let dates = 0;
  STAGES.forEach(k => { o[k] = date_(s[k], k); if (o[k]) dates++; });
  if (dates < 2) throw new Error('Вкажіть щонайменше дві дати етапів');
  LISTS.forEach(k => {
    const arr = Array.isArray(c[k]) ? c[k] : [];
    if (arr.length > MAX_LIST) throw new Error('Забагато дат у списку ' + k);
    o[k] = arr.map(d => date_(d, k)).filter(Boolean).sort().join(' ');
  });
  o.nick = String(c.nick || '').trim().slice(0, 40);
  o.stream = STREAMS.indexOf(c.stream) >= 0 ? c.stream : '';
  o.ee = ['ee', 'base'].indexOf(c.ee) >= 0 ? c.ee : '';
  o.security = ['yes', 'no'].indexOf(c.security) >= 0 ? c.security : '';
  o.notes = String(c.notes || '').trim().slice(0, 500);
  return o;
}

function writeRow_(row, obj) {
  const sh = sheet_();
  const h = headers_(sh);
  const rg = sh.getRange(row, 1, 1, h.length);
  rg.setNumberFormat('@');
  rg.setValues([h.map(n => obj[n] == null ? '' : obj[n])]);
}

function create_(c) {
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get('creates') || 0);
  if (n >= MAX_CREATES_PER_10_MIN) throw new Error('Забагато нових кейсів за короткий час, спробуйте пізніше');
  cache.put('creates', String(n + 1), 600);

  const o = clean_(c);
  o.id = Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  o.key = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  o.created = o.updated = new Date().toISOString();
  writeRow_(sheet_().getLastRow() + 1, o);
  return { ok: true, id: o.id, key: o.key };
}

function find_(id, key) {
  const r = rows_().find(o => String(o.id) === String(id));
  if (!r || String(r.key) !== String(key)) throw new Error('Кейс не знайдено або код редагування неправильний');
  return r;
}

function update_(id, key, c) {
  const r = find_(id, key);
  const o = clean_(c);
  o.id = String(r.id); o.key = String(r.key); o.created = fmt_(r.created); o.updated = new Date().toISOString();
  writeRow_(r._row, o);
  return { ok: true, id: o.id };
}

function delete_(id, key) {
  const r = find_(id, key);
  sheet_().deleteRow(r._row);
  return { ok: true };
}
