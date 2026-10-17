/**
 * Бекенд дашборду PR-таймлайнів.
 * Вставте цей код у Google Таблицю → Розширення → Apps Script,
 * потім Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone).
 */

const SHEET_NAME = 'cases';
const STAGES = ['eoi','laa','mpnp','verif','nom','pr','aor','bio','pal','med','fd','p1','p2','ecopr'];
const META = ['nick','stream','ee','security','notes'];
const HEADERS = ['id','key','created','updated'].concat(META, STAGES);
const STREAMS = ['swm','swo','ies','bis','other'];
const MAX_CREATES_PER_10_MIN = 30;

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
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

function clean_(c) {
  if (!c || typeof c !== 'object') throw new Error('Порожні дані');
  const s = c.s || {};
  let dates = 0;
  const stageVals = STAGES.map(k => {
    const d = String(s[k] || '').trim();
    if (!d) return '';
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) throw new Error('Неправильна дата: ' + k);
    const y = Number(d.slice(0, 4));
    if (y < 2015 || y > 2035) throw new Error('Неправдоподібний рік: ' + d);
    dates++;
    return d;
  });
  if (dates < 2) throw new Error('Вкажіть щонайменше дві дати етапів');
  const nick = String(c.nick || '').trim().slice(0, 40);
  const stream = STREAMS.indexOf(c.stream) >= 0 ? c.stream : '';
  const ee = ['ee', 'base'].indexOf(c.ee) >= 0 ? c.ee : '';
  const security = ['yes', 'no'].indexOf(c.security) >= 0 ? c.security : '';
  const notes = String(c.notes || '').trim().slice(0, 500);
  return [nick, stream, ee, security, notes].concat(stageVals);
}

function writeRow_(row, values) {
  const rg = sheet_().getRange(row, 1, 1, HEADERS.length);
  rg.setNumberFormat('@');
  rg.setValues([values]);
}

function create_(c) {
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get('creates') || 0);
  if (n >= MAX_CREATES_PER_10_MIN) throw new Error('Забагато нових кейсів за короткий час, спробуйте пізніше');
  cache.put('creates', String(n + 1), 600);

  const vals = clean_(c);
  const id = Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  const key = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  const now = new Date().toISOString();
  writeRow_(sheet_().getLastRow() + 1, [id, key, now, now].concat(vals));
  return { ok: true, id: id, key: key };
}

function find_(id, key) {
  const r = rows_().find(o => String(o.id) === String(id));
  if (!r || String(r.key) !== String(key)) throw new Error('Кейс не знайдено або код редагування неправильний');
  return r;
}

function update_(id, key, c) {
  const r = find_(id, key);
  const vals = clean_(c);
  writeRow_(r._row, [r.id, r.key, fmt_(r.created), new Date().toISOString()].concat(vals));
  return { ok: true, id: String(r.id) };
}

function delete_(id, key) {
  const r = find_(id, key);
  sheet_().deleteRow(r._row);
  return { ok: true };
}
