'use strict';
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const zlib = require('zlib');

const core = require('../lib/core');
const { readZip, writeZip, crc32 } = require('../lib/zip');
const { openWorkbook, formatNumber } = require('../lib/xlsx');
const { TrueTypeFont } = require('../lib/ttf');
const { decodePng } = require('../lib/image');
const { renderPdf, findPlaceholders, findLabeled, wrapLines, fontSafe, layoutSheet } = require('../lib/render');

const FONT_DIR = path.join(__dirname, '..', 'fonts');
const fonts = {
  regular: new TrueTypeFont(fs.readFileSync(path.join(FONT_DIR, 'NanumMyeongjo-Regular.ttf'))),
  bold: new TrueTypeFont(fs.readFileSync(path.join(FONT_DIR, 'NanumMyeongjo-Bold.ttf'))),
};

/* ───────────── 테스트용 엑셀·PNG 만들기 ───────────── */

function png(w, h, rgba) {
  const raw = Buffer.alloc((w * 4 + 1) * h);
  for (let y = 0; y < h; y++) rgba.copy(raw, y * (w * 4 + 1) + 1, y * w * 4, (y + 1) * w * 4);
  const chunk = (type, data) => {
    const body = Buffer.concat([Buffer.from(type, 'latin1'), data]);
    const len = Buffer.alloc(4); len.writeUInt32BE(data.length);
    const crc = Buffer.alloc(4); crc.writeUInt32BE(crc32(body));
    return Buffer.concat([len, body, crc]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(w, 0); ihdr.writeUInt32BE(h, 4); ihdr[8] = 8; ihdr[9] = 6;
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}

const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** 공문 양식 모양의 작은 엑셀 파일. cells: { B2: [값, 스타일번호] } */
function makeXlsx({ cells, merges = [], colWidth = 10, fit = false, image = true, extraSheet = false }) {
  const shared = [];
  const rows = {};
  for (const [ref, [value, s]] of Object.entries(cells)) {
    const r = Number(ref.replace(/^[A-Z]+/, ''));
    (rows[r] = rows[r] || []).push(`<c r="${ref}" s="${s || 0}" t="s"><v>${shared.push(value) - 1}</v></c>`);
  }
  const sheetData = Object.keys(rows).sort((a, b) => a - b).map(r => `<row r="${r}" ht="24" customHeight="1">${rows[r].join('')}</row>`).join('');
  const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
${fit ? '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' : ''}<sheetFormatPr defaultRowHeight="15"/>
<cols><col min="1" max="6" width="${colWidth}" customWidth="1"/></cols>
<sheetData>${sheetData}</sheetData>
${merges.length ? `<mergeCells count="${merges.length}">${merges.map(m => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>` : ''}
<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
<pageSetup paperSize="9" orientation="portrait" ${fit ? 'fitToHeight="0"' : ''}/>
${image ? '<drawing r:id="rId1"/>' : ''}
</worksheet>`;
  const styles = `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2"><font><sz val="11"/><name val="맑은 고딕"/></font><font><b/><sz val="16"/><color rgb="FF000000"/><name val="Batang"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/></border><border><left/><right/><top/><bottom style="medium"><color rgb="FF000000"/></bottom></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0"><alignment horizontal="center" vertical="center"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="1"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment vertical="top" wrapText="1"/></xf></cellXfs>
</styleSheet>`;
  const files = [
    ['[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'],
    ['_rels/.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="officeDocument" Target="xl/workbook.xml"/></Relationships>'],
    ['xl/workbook.xml', `<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${extraSheet ? '<sheet name="안내" sheetId="2" r:id="rId9"/>' : ''}<sheet name="양식" sheetId="1" r:id="rId1"/></sheets></workbook>`],
    ['xl/_rels/workbook.xml.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId9" Type="worksheet" Target="worksheets/sheet2.xml"/></Relationships>'],
    ['xl/worksheets/sheet1.xml', sheet],
    ['xl/worksheets/sheet2.xml', '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>안내 시트</t></is></c></row></sheetData></worksheet>'],
    ['xl/sharedStrings.xml', `<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${shared.map(s => `<si><t xml:space="preserve">${esc(s)}</t></si>`).join('')}</sst>`],
    ['xl/styles.xml', styles],
  ];
  if (image) {
    files.push(
      ['xl/worksheets/_rels/sheet1.xml.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="drawing" Target="../drawings/drawing1.xml"/></Relationships>'],
      ['xl/drawings/drawing1.xml', `<?xml version="1.0"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>12700</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>12700</xdr:rowOff></xdr:from><xdr:ext cx="508000" cy="508000"/>
<xdr:pic><xdr:blipFill><a:blip r:embed="rId1"/><a:srcRect t="10000" b="10000"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>`],
      ['xl/drawings/_rels/drawing1.xml.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="image" Target="../media/image1.png"/></Relationships>'],
    );
  }
  const zipFiles = files.map(([name, xml]) => ({ name, data: Buffer.from(xml, 'utf8') }));
  if (image) zipFiles.push({ name: 'xl/media/image1.png', data: png(2, 2, Buffer.from([255, 0, 0, 255, 0, 255, 0, 255, 0, 0, 255, 128, 255, 255, 255, 0])) });
  return writeZip(zipFiles);
}

const FORM = {
  cells: {
    B2: ['수신자', 1], C2: ['[교회명] 담임목사님', 1],
    B3: ['제목', 1], C3: ['2026년 북한중보기도주간 후원 협조 요청', 1],
    B4: ['1. 주님의 이름으로 문안드립니다. 「2026 북한중보기도주간」을 진행하고자 합니다.', 3],
    B6: ['시행', 0], C6: ['북한중보기도주간 TFT - [문서번호]([날짜])', 2],
  },
  merges: ['C2:F2', 'C3:F3', 'B4:F5', 'C6:F6'],
};

/** PDF 안의 스트림을 모두 풀어서 돌려준다(이 작성기가 쓰는 형식만 읽는다) */
function pdfStreams(buf) {
  const s = buf.toString('latin1'), out = [];
  const re = /<< \/Length (\d+)((?:(?!>>\nstream\n)[\s\S])*)>>\nstream\n/g;
  let m;
  while ((m = re.exec(s))) {
    const data = buf.subarray(m.index + m[0].length, m.index + m[0].length + Number(m[1]));
    out.push({ dict: m[2], data: /FlateDecode/.test(m[2]) ? zlib.inflateSync(data) : data });
  }
  return out;
}
const hexOf = (ttf, str) => [...str].map(ch => ttf.glyphFor(ch.codePointAt(0)).toString(16).padStart(4, '0')).join('');

/* ───────────── 번호·파일 이름 ───────────── */

test('교회명 목록: 줄마다 한 곳, 겹친 이름은 한 번만', () => {
  assert.deepStrictEqual(core.parseChurches('한동교회\n\n 포항  중앙교회 \r\n한동교회'), { list: ['한동교회', '포항 중앙교회'], dup: ['한동교회'] });
  // 엑셀에서 두 열을 복사하면 첫 열만, 한 줄이면 탭·쉼표로 나눈다
  assert.deepStrictEqual(core.parseChurches('가교회\t홍길동\n나교회\t김철수').list, ['가교회', '나교회']);
  assert.deepStrictEqual(core.parseChurches('가교회\t나교회, 다교회').list, ['가교회', '나교회', '다교회']);
});

test('파일 이름: [제목]_[교회명] 형식, 못 쓰는 글자는 뺀다', () => {
  const s = core.defaultState().settings;
  const t = '2026년 북한중보기도주간 후원 협조 요청';
  assert.strictEqual(core.fileNameFor(s, { title: t, church: '한동교회' }), `[북한중보기도주간] ${t}_한동교회`);
  assert.strictEqual(core.fileNameFor(s, { title: 'a/b:c', church: '교회?' }), '[북한중보기도주간] a b c_교회');
  assert.strictEqual(core.fileNameFor({ fileNameFormat: '[문서번호]_[제목]' }, { title: '[교회명] 안내', church: '가교회', docNo: '007' }), '007_가교회 안내');
});

test('발급: 적은 순서대로 번호가 오르고 자릿수를 맞춘다', () => {
  const st = core.defaultState();
  st.form = { hash: 'h', sheet: '양식' };
  st.settings.nextNumber = 9; st.settings.digits = 3;
  const plan = core.plan(st, ['가교회', '나교회'], '2026.10.24', '제목');
  assert.deepStrictEqual(plan.map(p => p.docNo), ['009', '010']);
  assert.strictEqual(st.settings.nextNumber, 9, '계획만으로는 번호를 쓰지 않는다');
  const rows = core.issue(st, ['가교회', '나교회'], '2026.10.24', '제목');
  assert.deepStrictEqual(rows.map(r => [r.id, r.no, r.docNo, r.fileName]), [
    [1, 9, '009', '[북한중보기도주간] 제목_가교회.pdf'], [2, 10, '010', '[북한중보기도주간] 제목_나교회.pdf'],
  ]);
  assert.strictEqual(st.settings.nextNumber, 11);
  assert.deepStrictEqual(core.plan(st, ['가교회'], '2026.10.24', '제목')[0].issuedBefore, [{ no: 9, date: '2026.10.24' }]);
  assert.throws(() => core.issue(core.defaultState(), ['가교회'], '', ''), /양식/);
});

test('설정 검증과 이전 버전 데이터 옮기기', () => {
  const s = core.defaultState().settings;
  assert.throws(() => core.sanitizeSettings({ fileNameFormat: '[제목]' }, s, []), /\[교회명\]/);
  assert.throws(() => core.sanitizeSettings({ nextNumber: 0 }, s, []), /1 이상/);
  assert.match(core.sanitizeSettings({ nextNumber: 2 }, s, [{ no: 5 }]).warning, /겹칠/);
  const old = { settings: { org: '○○교회', nextNumber: 25, departments: [] }, ledger: [{ id: 1 }] };
  const m = core.migrate(old);
  assert.strictEqual(m.version, 2);
  assert.strictEqual(m.settings.nextNumber, 25);
  assert.deepStrictEqual(m.ledger, []);
  assert.strictEqual(core.dotDate('2026-10-24'), '2026.10.24');
  assert.throws(() => core.dotDate('24/10/2026'), /시행일자/);
  assert.strictEqual(core.kstDate(new Date('2026-10-23T15:30:00Z')).iso, '2026-10-24');
});

test('발행대장 CSV는 쉼표·따옴표를 감싸고 BOM을 붙인다', () => {
  const csv = core.ledgerCsv([{ docNo: '1', date: '2026.10.24', church: '가, 나', fileName: 'x.pdf', title: '"인용"', issuedAt: 't' }]);
  assert.ok(csv.startsWith('﻿문서번호'));
  assert.ok(csv.includes('"가, 나"') && csv.includes('"""인용"""'));
});

/* ───────────── ZIP·엑셀 읽기 ───────────── */

test('ZIP: 한글 파일 이름으로 쓰고 다시 읽는다', () => {
  const buf = writeZip([{ name: '[가] 공문_한동교회.pdf', data: Buffer.from('%PDF-1') }, { name: 'b.txt', data: Buffer.alloc(5000, 'a') }]);
  const z = readZip(buf);
  assert.deepStrictEqual(Object.keys(z), ['[가] 공문_한동교회.pdf', 'b.txt']);
  assert.strictEqual(z['[가] 공문_한동교회.pdf']().toString(), '%PDF-1');
  assert.strictEqual(z['b.txt']().length, 5000);
  assert.throws(() => readZip(Buffer.from('엑셀 아님')), /엑셀/);
});

test('엑셀 양식: 칸 값·병합·스타일·그림·자리 표시를 읽는다', () => {
  const wb = openWorkbook(makeXlsx({ ...FORM, extraSheet: true }));
  assert.deepStrictEqual(wb.sheets.map(s => s.name), ['안내', '양식']);
  const sh = wb.readSheet('양식');
  assert.strictEqual(sh.cells['2,3'].value, '[교회명] 담임목사님');
  assert.deepStrictEqual(sh.merges[0], { r1: 2, c1: 3, r2: 2, c2: 6 });
  assert.strictEqual(wb.styles.xfs[1].font.b, true);
  assert.strictEqual(wb.styles.xfs[2].border.bottom.style, 'medium');
  assert.strictEqual(sh.images.length, 1);
  assert.strictEqual(sh.images[0].type, 'png');
  assert.deepStrictEqual(sh.images[0].crop, { l: 0, t: 0.1, r: 0, b: 0.1 });
  assert.deepStrictEqual(findPlaceholders(sh), { 교회명: 1, 문서번호: 1, 날짜: 1 });
  assert.strictEqual(findLabeled(sh, '제목'), '2026년 북한중보기도주간 후원 협조 요청');
});

test('숫자·날짜 표시 형식', () => {
  assert.strictEqual(formatNumber(45954, 'yyyy.mm.dd'), '2025.10.24');
  assert.strictEqual(formatNumber(45954, 'yyyy"년" m"월" d"일"'), '2025년 10월 24일');
  assert.strictEqual(formatNumber(1234567.891, '#,##0.00'), '1,234,567.89');
  assert.strictEqual(formatNumber(0.25, '0%'), '25%');
  assert.strictEqual(formatNumber(3.14159265358979, 'General'), '3.141592654');
});

test('PNG 풀기: RGBA와 투명도', () => {
  const im = decodePng(png(2, 1, Buffer.from([10, 20, 30, 255, 40, 50, 60, 0])));
  assert.deepStrictEqual([im.width, im.height, [...im.rgb], [...im.alpha]], [2, 1, [10, 20, 30, 40, 50, 60], [255, 0]]);
});

/* ───────────── 글꼴·PDF ───────────── */

test('글꼴 부분 집합: 쓴 글리프만 남고 글리프 번호는 그대로', () => {
  const g = fonts.regular.glyphFor('한'.codePointAt(0));
  const other = fonts.regular.glyphFor('똠'.codePointAt(0));
  assert.ok(g > 0 && other > 0);
  const subBuf = fonts.regular.subset([g]);
  const sub = new TrueTypeFont(subBuf);
  assert.strictEqual(sub.numGlyphs, fonts.regular.numGlyphs);
  const [s1, e1] = sub.glyphRange(g), [s0, e0] = fonts.regular.glyphRange(g);
  const len = e0 - s0; // 부분 집합에서는 4바이트 단위로 맞추느라 뒤에 0이 붙을 수 있다
  assert.ok(e1 - s1 >= len && e1 - s1 < len + 4);
  const glyph = (f, s) => f.buf.subarray(f.tables.glyf.off + s, f.tables.glyf.off + s + len);
  assert.ok(glyph(sub, s1).equals(glyph(fonts.regular, s0)), '글리프 모양 데이터가 그대로');
  const [s2, e2] = sub.glyphRange(other);
  assert.strictEqual(e2 - s2, 0, '쓰지 않은 글리프는 비운다');
  assert.ok(subBuf.length < 300000, '3MB 글꼴이 작아진다');
});

test('줄바꿈과 글꼴에 없는 글자 처리', () => {
  const measure = s => s.length;
  assert.deepStrictEqual(wrapLines('가나 다라마 바', measure, 5), ['가나', '다라마 바']);
  assert.deepStrictEqual(wrapLines('가나다라마바사', measure, 3), ['가나다', '라마바', '사']);
  assert.deepStrictEqual(wrapLines('첫 줄\n\n셋째', measure, 10), ['첫 줄', '', '셋째']);
  const missing = new Set();
  assert.strictEqual(fontSafe('① 10㎡ 漢', fonts.regular, missing), '(1) 10m² 漢');
  assert.deepStrictEqual([...missing], ['漢']);
});

test('PDF: [교회명]·[문서번호]·[날짜]가 교회마다 채워진다', () => {
  const wb = openWorkbook(makeXlsx(FORM));
  const sh = wb.readSheet('양식');
  const letters = [
    { vars: { 교회명: '한동교회', 문서번호: '7', 날짜: '2026.10.24' } },
    { vars: { 교회명: '포항중앙교회', 문서번호: '8', 날짜: '2026.10.24' } },
  ];
  const r = renderPdf(sh, wb.styles, fonts, letters, { title: '시험' });
  assert.ok(r.pdf.subarray(0, 8).toString('latin1').startsWith('%PDF-1.7'));
  assert.ok(r.pdf.toString('latin1').trimEnd().endsWith('%%EOF'));
  assert.strictEqual((r.pdf.toString('latin1').match(/\/Type \/Page /g) || []).length, 2, '교회마다 한 쪽');
  assert.deepStrictEqual(r.missing, []);
  const streams = pdfStreams(r.pdf);
  const content = streams.filter(s => s.data.includes('Tj')).map(s => s.data.toString('latin1')).join('\n');
  // 수신자 칸(굵게)에 교회 이름, 시행 칸에 번호와 날짜
  assert.ok(content.includes(hexOf(fonts.bold, '한동교회 담임목사님')));
  assert.ok(content.includes(hexOf(fonts.bold, '포항중앙교회 담임목사님')));
  assert.ok(content.includes(hexOf(fonts.regular, '북한중보기도주간 TFT - 7(2026.10.24)')));
  assert.ok(!content.includes(hexOf(fonts.bold, '[교회명]')), '자리 표시가 남지 않는다');
  // 글자 복사·검색용 표에 한글이 들어 있다
  const cmaps = streams.filter(s => s.data.includes('beginbfchar')).map(s => s.data.toString('latin1')).join('');
  assert.ok(cmaps.includes('<d55c>') || cmaps.includes('<D55C>'));
  // 테두리(선)와 그림(투명도 포함)
  assert.match(content, / l S Q/);
  assert.ok(streams.some(s => /\/Subtype \/Image/.test(s.dict) && /\/SMask/.test(s.dict)));
});

test('한 페이지에 맞추기: 엑셀처럼 % 단위로 줄인다', () => {
  const wb = openWorkbook(makeXlsx({ ...FORM, colWidth: 30, fit: true, image: false }));
  const L = layoutSheet(wb.readSheet('양식'), wb.styles);
  // 내용이 있는 B~F열만 인쇄한다. 열 너비 30자(맑은 고딕 11pt 기준 240px=180pt) × 5열 = 900pt, 인쇄 폭 494.48pt → 54%
  assert.deepStrictEqual([L.range.c1, L.range.c2], [2, 6]);
  assert.strictEqual(L.scale, 0.54);
  assert.strictEqual(L.pages.length, 1);
});

/* ───────────── 서버 ───────────── */

test('서버: 양식 올리기 → 발급 → PDF·ZIP 내려받기', async () => {
  process.env.GONGMUN_DATA_DIR = fs.mkdtempSync(path.join(os.tmpdir(), 'gongmun-'));
  const { server, load } = require('../server');
  load();
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  const base = `http://127.0.0.1:${server.address().port}`;
  const json = (method, url, body) => fetch(base + url, { method, headers: { 'Content-Type': 'application/json' }, body: body && JSON.stringify(body) });
  try {
    let res = await fetch(base + '/api/preview', { method: 'POST', body: '{}' });
    assert.strictEqual(res.status, 400, '양식이 없으면 미리보기를 못 한다');

    res = await fetch(base + '/api/form', { method: 'POST', headers: { 'X-File-Name': encodeURIComponent('공문 양식.xlsx') }, body: makeXlsx(FORM) });
    const up = await res.json();
    assert.strictEqual(res.status, 200, JSON.stringify(up));
    assert.strictEqual(up.form.title, '2026년 북한중보기도주간 후원 협조 요청');
    assert.deepStrictEqual(up.missing, []);

    res = await json('POST', '/api/plan', { churches: '한동교회\n포항중앙교회\n한동교회', date: '2026-10-24' });
    const plan = await res.json();
    assert.deepStrictEqual(plan.rows.map(r => r.docNo), ['1', '2']);
    assert.deepStrictEqual(plan.duplicates, ['한동교회']);

    res = await json('POST', '/api/issue', { churches: '한동교회\n포항중앙교회', date: '2026-10-24' });
    const { records } = await res.json();
    assert.deepStrictEqual(records.map(r => r.fileName), [
      '[북한중보기도주간] 2026년 북한중보기도주간 후원 협조 요청_한동교회.pdf',
      '[북한중보기도주간] 2026년 북한중보기도주간 후원 협조 요청_포항중앙교회.pdf',
    ]);

    res = await fetch(`${base}/api/download?ids=${records[0].id}`);
    assert.strictEqual(res.headers.get('content-type'), 'application/pdf');
    assert.ok(res.headers.get('content-disposition').includes("filename*=UTF-8''" + encodeURIComponent(records[0].fileName).replace(/['()*]/g, c => '%' + c.charCodeAt(0).toString(16).toUpperCase())));
    assert.ok(Buffer.from(await res.arrayBuffer()).toString('latin1', 0, 5) === '%PDF-');

    res = await fetch(`${base}/api/download?ids=${records.map(r => r.id).join(',')}`);
    assert.strictEqual(res.headers.get('content-type'), 'application/zip');
    const zip = readZip(Buffer.from(await res.arrayBuffer()));
    assert.deepStrictEqual(Object.keys(zip), records.map(r => r.fileName));
    assert.strictEqual(zip[records[1].fileName]().toString('latin1', 0, 5), '%PDF-');

    res = await fetch(`${base}/api/download?ids=${records.map(r => r.id).join(',')}&merge=1`);
    assert.strictEqual(res.headers.get('content-type'), 'application/pdf');

    res = await json('PUT', '/api/settings', { fileNameFormat: '[제목]만' });
    assert.strictEqual(res.status, 400);
    res = await json('PUT', '/api/settings', { fileNameFormat: '[교회명]_[문서번호]', digits: 3 });
    assert.strictEqual(res.status, 200);
    res = await json('POST', '/api/plan', { churches: '새교회' });
    assert.strictEqual((await res.json()).rows[0].fileName, '새교회_003.pdf');

    res = await fetch(base + '/api/ledger.csv');
    assert.match(await res.text(), /포항중앙교회/);

    res = await fetch(base + '/api/form', { method: 'POST', body: '엑셀 아님' });
    assert.strictEqual(res.status, 400);
  } finally {
    await new Promise(resolve => server.close(resolve));
    fs.rmSync(process.env.GONGMUN_DATA_DIR, { recursive: true, force: true });
  }
});

test('발행대장 삭제: 마지막 번호를 지우면 다음 번호가 되돌아간다', () => {
  const st = core.defaultState();
  st.form = { hash: 'h', sheet: '양식' };
  st.settings.nextNumber = 24;
  core.issue(st, ['가교회', '나교회', '다교회'], '2026.10.24', '제목'); // 24, 25, 26
  assert.deepStrictEqual(core.removeRecords(st, [2]), { removed: 1, nextNumber: 27 }, '중간 번호는 다음 번호를 바꾸지 않는다');
  assert.deepStrictEqual(core.removeRecords(st, [1, 3]), { removed: 2, nextNumber: 24 }, '모두 지우면 처음 번호로');
  core.issue(st, ['라교회', '마교회'], '2026.10.24', '제목'); // 24, 25
  core.removeRecords(st, [st.ledger[0].id]);
  const [r] = core.issue(st, ['바교회'], '2026.10.24', '제목');
  assert.strictEqual(r.docNo, '26');
  assert.strictEqual(new Set(st.ledger.map(x => x.id)).size, st.ledger.length, '내부 번호는 겹치지 않는다');
  assert.throws(() => core.removeRecords(st, [99]), /찾지 못/);
});
