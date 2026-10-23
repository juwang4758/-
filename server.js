'use strict';
/**
 * 공문 자동 작성 서비스 — 서버 (외부 패키지 없음)
 *   node server.js   → http://localhost:3100
 *
 * 엑셀 공문 양식을 한 번 올려 두면, 교회 이름만 적어 [교회명]·[문서번호]가 채워진 PDF를 내려받는다.
 * 여러 교회를 한 번에 적으면 교회마다 PDF를 만들어 ZIP 하나로 준다.
 *
 * 환경변수: PORT, GONGMUN_DATA_DIR(기본 data/), GONGMUN_PASSWORD(설정하면 접속 시 비밀번호 요구; 아이디는 아무거나)
 */
const http = require('http');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const core = require('./lib/core');
const { openWorkbook } = require('./lib/xlsx');
const { TrueTypeFont } = require('./lib/ttf');
const { renderPdf, findPlaceholders, findLabeled } = require('./lib/render');
const { writeZip } = require('./lib/zip');

const PORT = Number(process.env.PORT) || 3100;
const DATA_DIR = path.resolve(process.env.GONGMUN_DATA_DIR
  || (process.env.RAILWAY_VOLUME_MOUNT_PATH && path.join(process.env.RAILWAY_VOLUME_MOUNT_PATH, 'gongmun'))
  || path.join(__dirname, 'data'));
const DATA_FILE = path.join(DATA_DIR, 'gongmun.json');
const FORMS_DIR = path.join(DATA_DIR, 'forms');
const PASSWORD = process.env.GONGMUN_PASSWORD || '';
const INDEX = path.join(__dirname, 'public', 'index.html');
const MAX_UPLOAD = 15 * 1024 * 1024;

function fail(status, message) { throw Object.assign(new Error(message), { status }); }

/* ───────────── 저장 ───────────── */

let state;
function load() {
  try { state = core.migrate(JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'))); }
  catch (e) {
    if (e.code !== 'ENOENT') fs.copyFileSync(DATA_FILE, `${DATA_FILE}.broken-${Date.now()}`);
    state = core.defaultState();
  }
}
function save() {
  fs.mkdirSync(DATA_DIR, { recursive: true });
  fs.writeFileSync(`${DATA_FILE}.tmp`, JSON.stringify(state));
  fs.renameSync(`${DATA_FILE}.tmp`, DATA_FILE);
}

/* ───────────── 양식·글꼴 ───────────── */

let fonts = null;
function getFonts() {
  if (!fonts) {
    fonts = {
      regular: new TrueTypeFont(fs.readFileSync(path.join(__dirname, 'fonts', 'NanumMyeongjo-Regular.ttf'))),
      bold: new TrueTypeFont(fs.readFileSync(path.join(__dirname, 'fonts', 'NanumMyeongjo-Bold.ttf'))),
    };
  }
  return fonts;
}

const workbooks = new Map(); // 양식 해시 → { wb, sheets: Map }
function getSheet(hash, sheetName) {
  let entry = workbooks.get(hash);
  if (!entry) {
    let buf;
    try { buf = fs.readFileSync(path.join(FORMS_DIR, `${hash}.xlsx`)); }
    catch { fail(410, '이 공문을 만들 때 쓴 양식 파일이 없습니다.'); }
    entry = { wb: openWorkbook(buf), sheets: new Map() };
    workbooks.set(hash, entry);
  }
  const name = entry.wb.sheets.some(s => s.name === sheetName) ? sheetName : entry.wb.sheets[0].name;
  if (!entry.sheets.has(name)) entry.sheets.set(name, entry.wb.readSheet(name));
  return { wb: entry.wb, sheet: entry.sheets.get(name) };
}

/** [교회명]·[문서번호]가 들어 있는 시트를 고른다 */
function pickSheet(wb) {
  const visible = wb.sheets.filter(s => !s.hidden);
  const list = visible.length ? visible : wb.sheets;
  for (const s of list) {
    const found = findPlaceholders(wb.readSheet(s.name));
    if (found['교회명'] || found['문서번호']) return s.name;
  }
  return (list.find(s => /양식|공문|^B$/.test(s.name)) || list[0]).name;
}

function formInfo() {
  const f = state.form;
  if (!f) return null;
  let wb, sheet;
  try { ({ wb, sheet } = getSheet(f.hash, f.sheet)); }
  catch (e) { console.error('양식 파일을 읽지 못했습니다:', e.message); return null; } // 다시 올리도록 "양식 없음"으로 보인다
  return {
    name: f.name, uploadedAt: f.uploadedAt, sheet: sheet.name,
    sheets: wb.sheets.filter(s => !s.hidden).map(s => s.name),
    title: findLabeled(sheet, '제목'), placeholders: findPlaceholders(sheet),
  };
}

function renderRecords(records, title) {
  const { wb, sheet } = getSheet(records[0].form, records[0].sheet);
  const letters = records.map(r => ({ vars: { 교회명: r.church, 문서번호: r.docNo, 날짜: r.date } }));
  return renderPdf(sheet, wb.styles, getFonts(), letters, { title });
}

/* ───────────── HTTP ───────────── */

function authorized(req) {
  if (!PASSWORD) return true;
  const m = /^Basic (.+)$/.exec(req.headers.authorization || '');
  if (!m) return false;
  const given = Buffer.from(m[1], 'base64').toString().split(':').slice(1).join(':');
  const a = crypto.createHash('sha256').update(given).digest(), b = crypto.createHash('sha256').update(PASSWORD).digest();
  return crypto.timingSafeEqual(a, b);
}

function send(res, status, body, type = 'application/json; charset=utf-8', extra = {}) {
  res.writeHead(status, { 'Content-Type': type, 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff', ...extra });
  res.end(typeof body === 'string' || Buffer.isBuffer(body) ? body : JSON.stringify(body));
}

/** 한글 파일 이름이 그대로 저장되도록 RFC 5987 형식으로 보낸다 */
function disposition(kind, name) {
  const ext = path.extname(name) || '';
  const encoded = encodeURIComponent(name).replace(/['()*]/g, c => '%' + c.charCodeAt(0).toString(16).toUpperCase());
  return `${kind}; filename="gongmun${ext}"; filename*=UTF-8''${encoded}`;
}

function sendPdf(res, result, name, kind = 'attachment') {
  send(res, 200, result.pdf, 'application/pdf', {
    'Content-Disposition': disposition(kind, name),
    'X-Missing-Glyphs': encodeURIComponent(result.missing.join('')),
  });
}

function readRaw(req, limit) {
  return new Promise((resolve, reject) => {
    let size = 0; const chunks = [];
    req.on('data', c => {
      size += c.length;
      if (size > limit) { reject(Object.assign(new Error('파일이 너무 큽니다.'), { status: 413 })); req.destroy(); }
      else chunks.push(c);
    });
    req.on('end', () => resolve(Buffer.concat(chunks)));
    req.on('error', reject);
  });
}
async function readJson(req) {
  const raw = await readRaw(req, 1e6);
  try { return JSON.parse(raw.toString() || '{}'); } catch { fail(400, '잘못된 요청입니다.'); }
}

function requireForm() {
  const info = formInfo();
  if (!info) fail(400, '먼저 설정에서 공문 양식(엑셀 파일)을 올려 주세요.');
  return info;
}

const routes = {
  'GET /api/state': () => ({
    settings: state.settings, form: formInfo(), ledger: state.ledger, today: core.kstDate().iso,
    defaultFileFormat: core.DEFAULT_FILE_FORMAT,
  }),

  'POST /api/form': async req => {
    const buf = await readRaw(req, MAX_UPLOAD);
    let name = 'form.xlsx';
    try { name = decodeURIComponent(req.headers['x-file-name'] || name); } catch { /* 이름이 깨져도 계속 */ }
    const wb = openWorkbook(buf);
    const sheetName = pickSheet(wb);
    const sheet = wb.readSheet(sheetName);
    const hash = crypto.createHash('sha256').update(buf).digest('hex').slice(0, 24);
    // 미리 한 번 그려 보아 깨지는 양식이면 저장하지 않는다
    const test = renderPdf(sheet, wb.styles, getFonts(), [{ vars: { 교회명: '○○교회', 문서번호: '1', 날짜: core.kstDate().dot } }]);
    fs.mkdirSync(FORMS_DIR, { recursive: true });
    fs.writeFileSync(path.join(FORMS_DIR, `${hash}.xlsx`), buf);
    workbooks.delete(hash);
    state.form = { hash, name: path.basename(name).slice(0, 120), uploadedAt: new Date().toISOString(), sheet: sheetName };
    save();
    return { form: formInfo(), pages: test.pages, missing: test.missing };
  },

  'GET /api/form.xlsx': (req, res) => {
    if (!state.form) fail(404, '올린 양식이 없습니다.');
    const buf = fs.readFileSync(path.join(FORMS_DIR, `${state.form.hash}.xlsx`));
    send(res, 200, buf, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', { 'Content-Disposition': disposition('attachment', state.form.name) });
  },

  'PUT /api/settings': async req => {
    const body = await readJson(req);
    const { settings, warning } = core.sanitizeSettings(body, state.settings, state.ledger);
    if (body.sheet != null && state.form) {
      const { wb } = getSheet(state.form.hash, state.form.sheet);
      if (!wb.sheets.some(s => s.name === body.sheet)) fail(400, '양식에 없는 시트입니다.');
      state.form.sheet = body.sheet;
    }
    state.settings = settings;
    save();
    return { settings, form: formInfo(), warning };
  },

  'POST /api/plan': async req => {
    const body = await readJson(req);
    const form = state.form ? formInfo() : null;
    const { list, dup } = core.parseChurches(body.churches);
    const date = body.date ? core.dotDate(body.date) : core.kstDate().dot;
    return { rows: core.plan(state, list, date, form ? form.title : ''), duplicates: dup };
  },

  'POST /api/preview': async (req, res) => {
    const form = requireForm();
    const body = await readJson(req);
    const { list } = core.parseChurches(body.church);
    const date = body.date ? core.dotDate(body.date) : core.kstDate().dot;
    const p = core.plan(state, list.length ? list.slice(0, 1) : ['[교회명]'], date, form.title)[0];
    const result = renderRecords([{ form: state.form.hash, sheet: state.form.sheet, church: list[0] || '', docNo: p.docNo, date }], p.fileName.replace(/\.pdf$/, ''));
    sendPdf(res, result, p.fileName, 'inline');
  },

  'POST /api/issue': async req => {
    const form = requireForm();
    const body = await readJson(req);
    const { list } = core.parseChurches(body.churches);
    const date = body.date ? core.dotDate(body.date) : core.kstDate().dot;
    const records = core.issue(state, list, date, form.title);
    save();
    return { records };
  },

  'GET /api/download': (req, res, url) => {
    const ids = (url.searchParams.get('ids') || '').split(',').map(Number).filter(n => n > 0);
    const records = ids.map(id => state.ledger.find(r => r.id === id)).filter(Boolean);
    if (!records.length) fail(404, '내려받을 공문을 찾지 못했습니다.');
    if (records.length > 300) fail(400, '한 번에 300개까지 내려받을 수 있습니다.');
    if (records.length === 1) {
      const r = records[0];
      return sendPdf(res, renderRecords([r], r.fileName.replace(/\.pdf$/, '')), r.fileName);
    }
    const groupName = core.fileNameFor(state.settings, { title: records[0].title, church: `교회 ${records.length}곳`, docNo: '', date: records[0].date });
    const sameForm = records.every(r => r.form === records[0].form && r.sheet === records[0].sheet);
    if (url.searchParams.get('merge') === '1' && sameForm) {
      return sendPdf(res, renderRecords(records, groupName), `${groupName}.pdf`);
    }
    // 교회마다 PDF 하나씩, ZIP으로 묶는다(같은 이름이 있으면 뒤에 번호를 붙인다)
    const used = new Map();
    const files = records.map(r => {
      let name = r.fileName;
      const n = (used.get(name) || 0) + 1;
      used.set(r.fileName, n);
      if (n > 1) name = name.replace(/\.pdf$/, ` (${n}).pdf`);
      return { name, data: renderRecords([r], r.fileName.replace(/\.pdf$/, '')).pdf };
    });
    send(res, 200, writeZip(files), 'application/zip', { 'Content-Disposition': disposition('attachment', `${groupName}.zip`) });
  },

  'POST /api/remove': async req => {
    const body = await readJson(req);
    const result = core.removeRecords(state, Array.isArray(body.ids) ? body.ids : []);
    save();
    return result;
  },

  'GET /api/ledger.csv': (req, res) => {
    send(res, 200, core.ledgerCsv(state.ledger), 'text/csv; charset=utf-8', { 'Content-Disposition': disposition('attachment', '공문 발행대장.csv') });
  },
};

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url, 'http://x');
    if (url.pathname === '/healthz') return send(res, 200, 'ok', 'text/plain');
    if (!authorized(req)) return send(res, 401, '비밀번호가 필요합니다.', 'text/plain; charset=utf-8', { 'WWW-Authenticate': 'Basic realm="gongmun", charset="UTF-8"' });
    if (req.method === 'GET' && url.pathname === '/') return send(res, 200, fs.readFileSync(INDEX), 'text/html; charset=utf-8');
    const route = routes[`${req.method} ${url.pathname}`];
    if (!route) return send(res, 404, { error: '없는 주소입니다.' });
    const out = await route(req, res, url);
    if (out !== undefined && !res.headersSent) send(res, 200, out);
  } catch (e) {
    if (res.headersSent) { res.destroy(); return; }
    send(res, e.status || 500, { error: e.status ? e.message : '처리하지 못했습니다. 양식 파일을 확인해 주세요.' });
    if (!e.status) console.error(e);
  }
});

if (require.main === module) {
  load();
  server.listen(PORT, () => console.log(`공문 자동 작성 서비스: http://localhost:${PORT}${PASSWORD ? ' (비밀번호 사용)' : ' (비밀번호 없음 — GONGMUN_PASSWORD 설정 권장)'}`));
}
module.exports = { server, load: () => load(), get state() { return state; } };
