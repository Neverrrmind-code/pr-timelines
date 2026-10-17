/**
 * Бекенд дашборду PR-таймлайнів.
 * Вставте цей код у Google Таблицю → Розширення → Apps Script,
 * потім Deploy → New deployment → Web app (Execute as: Me, Who has access: Anyone).
 * Після змін коду: Deploy → Manage deployments → ✏️ → Version: New version → Deploy.
 *
 * Для імпорту з Telegram потрібні Script Properties (⚙ Project Settings → Script properties):
 *   ADMIN_PASSWORD  — пароль сторінки імпорту
 *   GEMINI_API_KEY  — ключ з https://aistudio.google.com/apikey
 *   GEMINI_MODEL    — (необов'язково) назва моделі, за замовчуванням gemini-2.5-flash
 */

const SHEET_NAME = 'cases';
const STAGES = ['eoi','laa','mpnp','verif','nom','pr','aor','bio','sec','secdone','pal','med','fd','p1','p2','ecopr'];
const META = ['nick','stream','ee','security','notes','source'];
const LISTS = ['adr','ghost'];           // багаторазові події, дати через пробіл
const PRIVATE = ['srcref'];              // не віддається назовні (посилання на джерело)
const HEADERS = ['id','key','created','updated'].concat(META, STAGES, LISTS, PRIVATE);
const STREAMS = ['swm','swo','ies','bis','other'];
const MAX_CREATES_PER_10_MIN = 30;
const MAX_LIST = 20;
const MAX_IMPORT = 100;
const DEFAULT_MODEL = 'gemini-2.5-flash';

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

function prop_(k) {
  return String(PropertiesService.getScriptProperties().getProperty(k) || '').trim();
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

    // адмін-дії без запису: не тримаємо блокування, бо виклик ШІ може тривати довго
    if (b.action === 'adminCheck') {
      checkAdmin_(b.password);
      const refs = rows_().map(r => fmt_(r.srcref)).filter(Boolean);
      return json_({ ok: true, model: prop_('GEMINI_MODEL') || DEFAULT_MODEL, hasKey: !!prop_('GEMINI_API_KEY'), refs: refs });
    }
    if (b.action === 'parse') { checkAdmin_(b.password); return json_(parse_(b.text, b.today)); }

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      if (b.action === 'create') return json_(create_(b.case));
      if (b.action === 'update') return json_(update_(b.id, b.key, b.case));
      if (b.action === 'delete') return json_(delete_(b.id, b.key));
      if (b.action === 'import') { checkAdmin_(b.password); return json_(import_(b.cases)); }
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

function newIds_(o) {
  o.id = Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  o.key = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  o.created = o.updated = new Date().toISOString();
}

function create_(c) {
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get('creates') || 0);
  if (n >= MAX_CREATES_PER_10_MIN) throw new Error('Забагато нових кейсів за короткий час, спробуйте пізніше');
  cache.put('creates', String(n + 1), 600);

  const o = clean_(c);
  o.source = 'form';
  newIds_(o);
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
  o.source = fmt_(r.source) || 'form'; o.srcref = fmt_(r.srcref);
  writeRow_(r._row, o);
  return { ok: true, id: o.id };
}

function delete_(id, key) {
  const r = find_(id, key);
  sheet_().deleteRow(r._row);
  return { ok: true };
}

/* ===================== Адмін: імпорт з Telegram ===================== */

function checkAdmin_(pw) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('adminfail') || 0);
  if (fails >= 5) throw new Error('Забагато невдалих спроб входу. Зачекайте 10 хвилин.');
  const real = prop_('ADMIN_PASSWORD');
  if (!real) throw new Error('Пароль адміністратора не налаштовано (ADMIN_PASSWORD у Script Properties).');
  if (String(pw || '') !== real) {
    cache.put('adminfail', String(fails + 1), 600);
    throw new Error('Неправильний пароль');
  }
}

function import_(cases) {
  if (!Array.isArray(cases) || !cases.length) throw new Error('Немає кейсів для імпорту');
  if (cases.length > MAX_IMPORT) throw new Error('За один раз можна імпортувати до ' + MAX_IMPORT + ' кейсів');
  const sh = sheet_();
  const seen = {};
  rows_().forEach(r => { const ref = fmt_(r.srcref); if (ref) seen[ref] = true; });
  const added = [], skipped = [], errors = [];
  cases.forEach((c, i) => {
    try {
      const ref = String(c && c.srcref || '').trim().slice(0, 200);
      if (ref && seen[ref]) { skipped.push(i); return; }
      const o = clean_(c);
      o.source = 'telegram';
      o.srcref = ref;
      newIds_(o);
      writeRow_(sh.getLastRow() + 1, o);
      if (ref) seen[ref] = true;
      added.push({ i: i, id: o.id });
    } catch (err) {
      errors.push({ i: i, error: String(err && err.message || err) });
    }
  });
  return { ok: true, added: added, skipped: skipped, errors: errors };
}

function parse_(text, today) {
  const key = prop_('GEMINI_API_KEY');
  if (!key) throw new Error('Ключ Gemini не налаштовано (GEMINI_API_KEY у Script Properties).');
  const model = prop_('GEMINI_MODEL') || DEFAULT_MODEL;
  text = String(text || '').slice(0, 60000);
  if (!text.trim()) throw new Error('Порожній текст');
  today = /^\d{4}-\d{2}-\d{2}$/.test(String(today || '')) ? today : Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd');

  const body = {
    systemInstruction: { parts: [{ text: PROMPT_ }] },
    contents: [{ role: 'user', parts: [{ text: 'TODAY: ' + today + '\n\nMESSAGES:\n' + text }] }],
    generationConfig: { temperature: 0, responseMimeType: 'application/json' }
  };
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent';
  const r = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-goog-api-key': key }, payload: JSON.stringify(body)
  });
  const code = r.getResponseCode();
  let j;
  try { j = JSON.parse(r.getContentText()); } catch (e) { throw new Error('Gemini відповів не JSON (HTTP ' + code + ')'); }
  if (code !== 200) throw new Error('Gemini: ' + ((j.error && j.error.message) || ('HTTP ' + code)));
  const cand = (j.candidates || [])[0] || {};
  const parts = (cand.content && cand.content.parts) || [];
  const out = parts.map(p => p.text || '').join('').trim();
  if (!out) throw new Error('Gemini повернув порожню відповідь' + (cand.finishReason ? ' (' + cand.finishReason + ')' : ''));
  let data;
  try { data = JSON.parse(out.replace(/^```(?:json)?\s*|\s*```$/g, '')); }
  catch (e) { throw new Error('Не вдалося прочитати відповідь Gemini як JSON'); }
  const items = Array.isArray(data) ? data : (data.items || []);
  return { ok: true, model: model, items: items };
}

const PROMPT_ = [
  'You extract Canadian immigration (Manitoba MPNP -> IRCC permanent residence) timelines from Telegram messages written in Ukrainian, Russian or English.',
  'Each message in the input starts with a header line like "### MSG ref=<id> date=<YYYY-MM-DD>". Pasted text may have no headers; then treat the whole text as messages separated by blank lines.',
  '',
  'Return ONLY JSON: {"items":[ ... ]}. One item per distinct applicant timeline that contains at least 2 dated milestones. Skip questions, congratulations, chit-chat and messages without dated milestones.',
  'Item fields:',
  '- "s": object of milestone dates, ISO "YYYY-MM-DD", only for milestones explicitly dated in the text. Keys:',
  '  eoi = MPNP EOI / Expression of Interest submitted',
  '  laa = LAA, Letter of Advice to Apply, ITA from MPNP',
  '  mpnp = MPNP application submitted',
  '  verif = employment verification / phone call or email from MPNP checking the job',
  '  nom = MPNP nomination / approval letter / номінація',
  '  pr = PR application submitted to IRCC (подача на PR, IRCC submitted)',
  '  aor = AOR, Acknowledgement of Receipt',
  '  bio = biometrics given/completed (not the date BIL letter was received)',
  '  med = medicals passed/cleared/"medical exam passed" status (not the date the exam was attended)',
  '  sec = background/security check started (in progress)',
  '  secdone = background/security check completed',
  '  pal = PAL, Invitation to pre-arrival services',
  '  fd = final decision / decision made / approved',
  '  p1 = Portal 1 (first PR portal email, confirm presence in Canada)',
  '  p2 = Portal 2 (second portal email, photo upload / eCoPR request)',
  '  ecopr = eCoPR / COPR received',
  '- "adr": array of ISO dates of Additional Documents Requests (ADR).',
  '- "ghost": array of ISO dates of "ghost updates" (tracker changed without visible reason).',
  '- "stream": "swm" (Skilled Worker in Manitoba), "swo" (Skilled Worker Overseas), "ies" (International Education), "bis" (Business Investor), "other", or "" if unknown.',
  '- "ee": "ee" if Express Entry, "base" if paper/non-EE/base application, "" if unknown.',
  '- "security": "yes" if a comprehensive security/background check is mentioned, "no" if explicitly standard, else "".',
  '- "nick": short anonymous description in the language of the message, e.g. "Сім\'я з 2, retail supervisor". NEVER include personal names, @usernames, phone numbers, UCI or application numbers.',
  '- "notes": up to 300 chars of useful context (language test scores, EOI points, NOC/job, family size, job change, re-nomination). No personal names or identifiers.',
  '- "ambiguous": array of milestone keys (or "adr"/"ghost") whose day/month order could not be determined with certainty.',
  '- "msgref": the ref value from the message header ("" if none).',
  '- "msgdate": the date value from the message header ("" if none).',
  '- "quote": the first 160 characters of the source message, verbatim.',
  '',
  'Date rules:',
  '- Determine the date format per message. If any date in the message has a first number > 12, the message uses DD/MM; if any has a second number > 12, it uses MM/DD. Dots (12.05.2025) usually mean DD.MM. Slashes in English text are often MM/DD.',
  '- If still unclear, choose the interpretation that makes the milestones chronological (EOI < LAA < MPNP < nomination < PR < AOR < ... < eCoPR) and list the affected keys in "ambiguous".',
  '- If the year is missing, infer it from chronology and the message date; never produce dates after TODAY or after the message date.',
  '- Never invent milestones or dates that are not in the text. If a message lists two submissions of EOI, use the one that led to LAA and mention the other in notes.',
  '- If a message is an update to an earlier timeline by the same person, still output it as its own item.'
].join('\n');
