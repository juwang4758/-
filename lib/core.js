'use strict';
/**
 * 공문 자동 작성 — 번호 규칙·발행대장
 *
 * 엑셀 양식의 [교회명]·[문서번호]·[날짜] 자리를 교회마다 채워 PDF를 만든다.
 *  - 문서번호는 발급할 때마다 1씩 오른다. 여러 교회를 한 번에 발급하면 적은 순서대로 번호를 받는다.
 *  - 파일 이름은 형식 문자열로 만든다. 기본값 "[북한중보기도주간] [제목]_[교회명]"
 */

const DEFAULT_FILE_FORMAT = '[북한중보기도주간] [제목]_[교회명]';
const MAX_CHURCHES = 300;

function defaultState() {
  return {
    version: 2,
    settings: { fileNameFormat: DEFAULT_FILE_FORMAT, nextNumber: 1, digits: 0 },
    form: null, // { hash, name, uploadedAt, sheet }
    ledger: [],
  };
}

/** 이전 버전(부서별 번호 발급) 데이터는 다음 문서번호만 이어받는다 */
function migrate(state) {
  if (state && state.version === 2) return state;
  const fresh = defaultState();
  const next = state && state.settings && Number(state.settings.nextNumber);
  if (next >= 1) fresh.settings.nextNumber = Math.floor(next);
  return fresh;
}

function bad(message) { return Object.assign(new Error(message), { status: 400 }); }

const pad = n => String(n).padStart(2, '0');

/** 한국 시간 기준 날짜 */
function kstDate(now = new Date()) {
  const k = new Date(now.getTime() + 9 * 3600 * 1000);
  const y = k.getUTCFullYear(), m = k.getUTCMonth() + 1, d = k.getUTCDate();
  return { iso: `${y}-${pad(m)}-${pad(d)}`, dot: `${y}.${pad(m)}.${pad(d)}` };
}

/** "2026-10-24" → "2026.10.24" */
function dotDate(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso || ''));
  if (!m) throw bad('시행일자를 확인해 주세요.');
  return `${m[1]}.${m[2]}.${m[3]}`;
}

/**
 * 한 줄에 한 교회. 겹치는 이름은 한 번만 쓴다.
 * 엑셀에서 여러 열을 복사해 붙여 넣으면(줄마다 탭이 있음) 줄마다 첫 칸만 교회명으로 본다.
 * 한 줄만 있으면 탭·쉼표로 나눈다(엑셀의 가로 한 줄 복사, "가교회, 나교회" 입력).
 */
function parseChurches(text) {
  const lines = String(text || '').split(/\r\n|\r|\n/).filter(l => l.trim());
  const cells = lines.length > 1 ? lines.map(l => l.split('\t')[0]) : (lines[0] || '').split(/[\t,]/);
  const seen = new Set(), list = [], dup = [];
  for (const raw of cells) {
    const name = raw.replace(/\s+/g, ' ').trim().slice(0, 60);
    if (!name) continue;
    if (seen.has(name)) { dup.push(name); continue; }
    seen.add(name);
    list.push(name);
  }
  if (list.length > MAX_CHURCHES) throw bad(`한 번에 ${MAX_CHURCHES}곳까지 만들 수 있습니다.`);
  return { list, dup };
}

function formatNo(n, digits) { return String(n).padStart(Math.min(6, Math.max(0, digits || 0)), '0'); }

/** 파일 이름에 쓸 수 없는 글자(\ / : * ? " < > |)를 빼고 너무 길면 줄인다 */
function safeFileName(name) {
  const s = String(name).normalize('NFC').replace(/[\\/:*?"<>|\u0000-\u001f]+/g, ' ').replace(/\s+/g, ' ').trim().replace(/^\.+/, '');
  return [...s].slice(0, 150).join('') || '공문';
}

/** 제목을 먼저 넣고 나머지 자리를 채운다(제목 칸에 [교회명]이 있어도 채워지도록) */
function fileNameFor(settings, { title, church, docNo, date }) {
  const vars = { 교회명: church || '', 문서번호: docNo || '', 날짜: date || '' };
  const name = (settings.fileNameFormat || DEFAULT_FILE_FORMAT)
    .replace(/\[\s*제목\s*\]/g, () => title || '')
    .replace(/\[\s*(교회명|문서번호|날짜|시행일자)\s*\]/g, (all, k) => vars[k === '시행일자' ? '날짜' : k]);
  return safeFileName(name);
}

/** 발급 전에 번호·파일 이름을 미리 계산한다(상태는 바꾸지 않음) */
function plan(state, churches, date, title) {
  const s = state.settings;
  return churches.map((church, i) => {
    const no = s.nextNumber + i, docNo = formatNo(no, s.digits);
    const before = state.ledger.filter(r => r.church === church);
    return {
      no, docNo, church, date,
      fileName: fileNameFor(s, { title, church, docNo, date }) + '.pdf',
      issuedBefore: before.map(r => ({ no: r.no, date: r.date })),
    };
  });
}

/** 번호를 확정하고 발행대장에 기록한다. state를 직접 바꾼다 */
function issue(state, churches, date, title, now = new Date()) {
  if (!state.form) throw bad('먼저 공문 양식(엑셀 파일)을 올려 주세요.');
  if (!churches.length) throw bad('교회명을 한 줄에 하나씩 입력해 주세요.');
  const rows = plan(state, churches, date, title).map((p, i) => ({
    id: state.ledger.length + i + 1, no: p.no, docNo: p.docNo, church: p.church, date: p.date, fileName: p.fileName, title,
    form: state.form.hash, sheet: state.form.sheet, issuedAt: now.toISOString(),
  }));
  state.ledger.push(...rows);
  state.settings.nextNumber += rows.length;
  return rows;
}

function sanitizeSettings(input, old, ledger) {
  const fileNameFormat = String(input.fileNameFormat == null ? old.fileNameFormat : input.fileNameFormat).trim().slice(0, 200) || DEFAULT_FILE_FORMAT;
  if (!/\[\s*교회명\s*\]/.test(fileNameFormat)) throw bad('파일 이름 형식에 [교회명]이 있어야 교회마다 다른 파일이 됩니다.');
  const next = Math.floor(Number(input.nextNumber == null ? old.nextNumber : input.nextNumber));
  if (!(next >= 1 && next < 1e7)) throw bad('다음 문서번호는 1 이상의 숫자여야 합니다.');
  const digits = Math.floor(Number(input.digits == null ? old.digits : input.digits));
  if (!(digits >= 0 && digits <= 6)) throw bad('자릿수는 0~6 사이로 정해 주세요.');
  const used = ledger.reduce((m, r) => Math.max(m, r.no), 0);
  return {
    settings: { fileNameFormat, nextNumber: next, digits },
    warning: next <= used ? `이미 ${used}번까지 발급했습니다. 다음 번호를 ${next}번으로 낮추면 번호가 겹칠 수 있습니다.` : '',
  };
}

function csvCell(v) { const s = String(v == null ? '' : v); return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s; }

/** 엑셀에서 바로 열리도록 UTF-8 BOM을 붙인 발행대장 CSV */
function ledgerCsv(ledger) {
  const head = ['문서번호', '시행일자', '교회명', '파일 이름', '제목', '발급 시각'];
  const rows = ledger.map(r => [r.docNo, r.date, r.church, r.fileName, r.title, r.issuedAt]);
  return '﻿' + [head, ...rows].map(r => r.map(csvCell).join(',')).join('\r\n') + '\r\n';
}

module.exports = {
  DEFAULT_FILE_FORMAT, defaultState, migrate, kstDate, dotDate, parseChurches, formatNo,
  safeFileName, fileNameFor, plan, issue, sanitizeSettings, ledgerCsv,
};
