'use strict';
/**
 * GitHub Pages(서버 없음)용: 서버의 /api/... 를 브라우저 안에서 처리한다.
 * 화면(index.html)은 서버판과 같고, fetch('/api/...')만 여기로 돌린다.
 * 설정·발행대장·양식은 이 브라우저의 localStorage에 저장한다.
 */
const core = require('../lib/core');
const { openWorkbook } = require('../lib/xlsx');
const { TrueTypeFont } = require('../lib/ttf');
const { renderPdf, findPlaceholders, findLabeled } = require('../lib/render');
const { writeZip } = require('../lib/zip');

const KEY = 'gongmun.state', FORM_KEY = h => `gongmun.form.${h}`;
const realFetch = window.fetch.bind(window);

let state;
try { state = core.migrate(JSON.parse(localStorage.getItem(KEY))); } catch { state = core.defaultState(); }
function save() { localStorage.setItem(KEY, JSON.stringify(state)); }

class HttpError extends Error { constructor(status, msg) { super(msg); this.status = status; } }
const fail = (s, m) => { throw new HttpError(s, m); };

let fontsP = null;
function getFonts() {
  if (!fontsP) {
    const load = async name => {
      const r = await realFetch(`fonts/${name}`);
      if (!r.ok) throw new HttpError(500, '글꼴 파일을 불러오지 못했습니다.');
      return new TrueTypeFont(Buffer.from(await r.arrayBuffer()));
    };
    fontsP = Promise.all([load('NanumMyeongjo-Regular.ttf'), load('NanumMyeongjo-Bold.ttf')]).then(([regular, bold]) => ({ regular, bold }));
    fontsP.catch(() => { fontsP = null; });
  }
  return fontsP;
}

const workbooks = new Map();
function getSheet(hash, sheetName) {
  let entry = workbooks.get(hash);
  if (!entry) {
    const b64 = localStorage.getItem(FORM_KEY(hash));
    if (!b64) fail(410, '이 공문을 만들 때 쓴 양식 파일이 없습니다.');
    entry = { wb: openWorkbook(Buffer.from(b64, 'base64')), sheets: new Map() };
    workbooks.set(hash, entry);
  }
  const name = entry.wb.sheets.some(s => s.name === sheetName) ? sheetName : entry.wb.sheets[0].name;
  if (!entry.sheets.has(name)) entry.sheets.set(name, entry.wb.readSheet(name));
  return { wb: entry.wb, sheet: entry.sheets.get(name) };
}
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
  try { ({ wb, sheet } = getSheet(f.hash, f.sheet)); } catch { return null; }
  return { name: f.name, uploadedAt: f.uploadedAt, sheet: sheet.name, sheets: wb.sheets.filter(s => !s.hidden).map(s => s.name),
    title: findLabeled(sheet, '제목'), placeholders: findPlaceholders(sheet) };
}
function requireForm() { const i = formInfo(); if (!i) fail(400, '먼저 설정에서 공문 양식(엑셀 파일)을 올려 주세요.'); return i; }
async function renderRecords(records, title) {
  const { wb, sheet } = getSheet(records[0].form, records[0].sheet);
  return renderPdf(sheet, wb.styles, await getFonts(), records.map(r => ({ vars: { 교회명: r.church, 문서번호: r.docNo, 날짜: r.date } })), { title });
}
function disposition(kind, name) {
  const dot = name.lastIndexOf('.'), ext = dot > 0 ? name.slice(dot) : '';
  const enc = encodeURIComponent(name).replace(/['()*]/g, c => '%' + c.charCodeAt(0).toString(16).toUpperCase());
  return `${kind}; filename="gongmun${ext}"; filename*=UTF-8''${enc}`;
}
const json = (obj, status = 200) => new Response(JSON.stringify(obj), { status, headers: { 'Content-Type': 'application/json; charset=utf-8' } });
const file = (data, type, headers) => new Response(new Blob([data], { type }), { status: 200, headers: { 'Content-Type': type, ...headers } });
const pdf = (result, name, kind = 'attachment') => file(result.pdf, 'application/pdf', { 'Content-Disposition': disposition(kind, name), 'X-Missing-Glyphs': encodeURIComponent(result.missing.join('')) });
async function body(init) {
  const b = init && init.body;
  if (b == null) return {};
  try { return JSON.parse(typeof b === 'string' ? b : await new Response(b).text()); } catch { fail(400, '잘못된 요청입니다.'); }
}

const routes = {
  'GET /api/state': () => json({ settings: state.settings, form: formInfo(), ledger: state.ledger, today: core.kstDate().iso, defaultFileFormat: core.DEFAULT_FILE_FORMAT }),
  'POST /api/form': async (url, init) => {
    const buf = Buffer.from(await new Response(init.body).arrayBuffer());
    if (buf.length > 15 * 1024 * 1024) fail(413, '파일이 너무 큽니다.');
    let name = 'form.xlsx';
    try { name = decodeURIComponent((init.headers || {})['X-File-Name'] || name); } catch { /* 그대로 */ }
    const wb = openWorkbook(buf);
    const sheetName = pickSheet(wb);
    const sheet = wb.readSheet(sheetName);
    const hash = require('crypto').createHash('sha256').update(buf).digest('hex').slice(0, 24);
    const test = renderPdf(sheet, wb.styles, await getFonts(), [{ vars: { 교회명: '○○교회', 문서번호: '1', 날짜: core.kstDate().dot } }]);
    try { localStorage.setItem(FORM_KEY(hash), buf.toString('base64')); }
    catch { fail(507, '브라우저 저장 공간이 부족합니다. 양식 파일 크기를 줄여 주세요.'); }
    workbooks.delete(hash);
    state.form = { hash, name: name.split(/[\\/]/).pop().slice(0, 120), uploadedAt: new Date().toISOString(), sheet: sheetName };
    save();
    return json({ form: formInfo(), pages: test.pages, missing: test.missing });
  },
  'GET /api/form.xlsx': () => {
    if (!state.form) fail(404, '올린 양식이 없습니다.');
    return file(Buffer.from(localStorage.getItem(FORM_KEY(state.form.hash)) || '', 'base64'), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', { 'Content-Disposition': disposition('attachment', state.form.name) });
  },
  'PUT /api/settings': async (url, init) => {
    const b = await body(init);
    const { settings, warning } = core.sanitizeSettings(b, state.settings, state.ledger);
    if (b.sheet != null && state.form) {
      const { wb } = getSheet(state.form.hash, state.form.sheet);
      if (!wb.sheets.some(s => s.name === b.sheet)) fail(400, '양식에 없는 시트입니다.');
      state.form.sheet = b.sheet;
    }
    state.settings = settings; save();
    return json({ settings, form: formInfo(), warning });
  },
  'POST /api/plan': async (url, init) => {
    const b = await body(init);
    const form = formInfo();
    const { list, dup } = core.parseChurches(b.churches);
    const date = b.date ? core.dotDate(b.date) : core.kstDate().dot;
    return json({ rows: core.plan(state, list, date, form ? form.title : ''), duplicates: dup });
  },
  'POST /api/preview': async (url, init) => {
    const form = requireForm();
    const b = await body(init);
    const { list } = core.parseChurches(b.church);
    const date = b.date ? core.dotDate(b.date) : core.kstDate().dot;
    const p = core.plan(state, list.length ? list.slice(0, 1) : ['[교회명]'], date, form.title)[0];
    const result = await renderRecords([{ form: state.form.hash, sheet: state.form.sheet, church: list[0] || '', docNo: p.docNo, date }], p.fileName.replace(/\.pdf$/, ''));
    return pdf(result, p.fileName, 'inline');
  },
  'POST /api/issue': async (url, init) => {
    const form = requireForm();
    const b = await body(init);
    const { list } = core.parseChurches(b.churches);
    const date = b.date ? core.dotDate(b.date) : core.kstDate().dot;
    const records = core.issue(state, list, date, form.title);
    save();
    return json({ records });
  },
  'GET /api/download': async url => {
    const ids = (url.searchParams.get('ids') || '').split(',').map(Number).filter(n => n > 0);
    const records = ids.map(id => state.ledger.find(r => r.id === id)).filter(Boolean);
    if (!records.length) fail(404, '내려받을 공문을 찾지 못했습니다.');
    if (records.length > 300) fail(400, '한 번에 300개까지 내려받을 수 있습니다.');
    if (records.length === 1) return pdf(await renderRecords(records, records[0].fileName.replace(/\.pdf$/, '')), records[0].fileName);
    const groupName = core.fileNameFor(state.settings, { title: records[0].title, church: `교회 ${records.length}곳`, docNo: '', date: records[0].date });
    if (url.searchParams.get('merge') === '1' && records.every(r => r.form === records[0].form && r.sheet === records[0].sheet)) {
      return pdf(await renderRecords(records, groupName), `${groupName}.pdf`);
    }
    const used = new Map(), files = [];
    for (const r of records) {
      const n = (used.get(r.fileName) || 0) + 1; used.set(r.fileName, n);
      files.push({ name: n > 1 ? r.fileName.replace(/\.pdf$/, ` (${n}).pdf`) : r.fileName, data: (await renderRecords([r], r.fileName.replace(/\.pdf$/, ''))).pdf });
    }
    return file(writeZip(files), 'application/zip', { 'Content-Disposition': disposition('attachment', `${groupName}.zip`) });
  },
  'POST /api/remove': async (url, init) => {
    const b = await body(init);
    const result = core.removeRecords(state, Array.isArray(b.ids) ? b.ids : []);
    save();
    return json(result);
  },
  'GET /api/ledger.csv': () => file(core.ledgerCsv(state.ledger), 'text/csv; charset=utf-8', { 'Content-Disposition': disposition('attachment', '공문 발행대장.csv') }),
};

window.fetch = async (input, init = {}) => {
  const url = new URL(typeof input === 'string' ? input : input.url, location.href);
  const idx = url.pathname.indexOf('/api/');
  if (idx < 0) return realFetch(input, init);
  const route = routes[`${(init.method || 'GET').toUpperCase()} ${url.pathname.slice(idx)}`];
  if (!route) return json({ error: '없는 주소입니다.' }, 404);
  try { return await route(url, init); }
  catch (e) {
    if (!(e instanceof HttpError) && !e.status) console.error(e);
    return json({ error: e.status ? e.message : '처리하지 못했습니다. 양식 파일을 확인해 주세요.' }, e.status || 500);
  }
};

document.addEventListener('DOMContentLoaded', () => {
  const note = document.createElement('div');
  note.style.cssText = 'background:#fff6e0;color:#7a4b00;padding:8px 16px;font-size:13px;text-align:center';
  note.textContent = '양식·문서번호·발행대장은 이 브라우저에만 저장됩니다. 다른 컴퓨터나 다른 브라우저와는 번호가 이어지지 않으니, 한 컴퓨터에서만 발급해 주세요.';
  document.body.insertBefore(note, document.body.children[1] || null);
});
