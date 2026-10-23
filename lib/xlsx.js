'use strict';
/**
 * 엑셀(.xlsx) 양식 읽기 (외부 패키지 없음)
 *
 * 공문 양식 시트를 PDF로 그리는 데 필요한 것만 읽는다.
 *  - 칸 값(공유 문자열·수식의 마지막 계산값 포함), 병합, 열 너비, 행 높이
 *  - 글꼴·정렬·줄바꿈·테두리·채우기
 *  - 인쇄 설정(여백·배율·한 페이지에 맞추기·인쇄 영역)
 *  - 그림(로고)과 위치, 자르기(srcRect)
 */
const path = require('path').posix;
const { readZip } = require('./zip');

/* ───────────── 아주 작은 XML 파서 ───────────── */

const ENT = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };
function decode(s) {
  return s.replace(/&(#x[0-9a-f]+|#\d+|lt|gt|amp|quot|apos);/gi, (all, e) => {
    if (e[0] !== '#') return ENT[e.toLowerCase()];
    const code = e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
    return code >= 0 && code <= 0x10ffff ? String.fromCodePoint(code) : '';
  });
}
/** OOXML 문자열 안의 _xHHHH_ 표기(줄바꿈 문자 등)를 실제 글자로 */
function unescapeOoxml(s) {
  return s.replace(/_x([0-9A-Fa-f]{4})_/g, (all, h) => String.fromCharCode(parseInt(h, 16)));
}

function parseXml(str) {
  const root = { name: '#root', attrs: {}, children: [], text: '' };
  const stack = [root];
  const re = /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[([\s\S]*?)\]\]>|<!DOCTYPE[^>]*>|<(\/?)([^\s/>]+)((?:\s+[^\s=]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>|([^<]+)/g;
  let m;
  while ((m = re.exec(str))) {
    const top = stack[stack.length - 1];
    if (m[1] != null) { top.text += m[1]; continue; }
    if (m[6] != null) { top.text += decode(m[6]); continue; }
    if (!m[3]) continue; // 주석·선언
    const local = m[3].includes(':') ? m[3].slice(m[3].indexOf(':') + 1) : m[3];
    if (m[2]) { // 닫는 태그
      if (stack.length > 1) stack.pop();
      continue;
    }
    const attrs = {};
    const ar = /([^\s=]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
    let a;
    while ((a = ar.exec(m[4]))) {
      attrs[a[1]] = decode(a[2] != null ? a[2] : a[3]);
    }
    const el = { name: local, attrs, children: [], text: '' };
    top.children.push(el);
    if (!m[5]) stack.push(el);
  }
  return root;
}
const kids = (el, name) => (el ? el.children.filter(c => c.name === name) : []);
const kid = (el, name) => (el ? el.children.find(c => c.name === name) || null : null);
function deep(el, name, out = []) {
  if (!el) return out;
  for (const c of el.children) { if (c.name === name) out.push(c); deep(c, name, out); }
  return out;
}
const attr = (el, k) => (el ? el.attrs[k] : undefined);
/** 접두사가 붙은 속성(r:id 등)을 접두사와 상관없이 찾는다 */
function attrLocal(el, local) {
  if (!el) return undefined;
  for (const k of Object.keys(el.attrs)) if (k === local || k.endsWith(':' + local)) return el.attrs[k];
  return undefined;
}

/* ───────────── 칸 주소 ───────────── */

function colIndex(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}
function parseRef(ref) {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(ref.trim());
  return m ? { c: colIndex(m[1]), r: Number(m[2]) } : null;
}
function parseRange(ref) {
  const [a, b] = ref.replace(/^.*!/, '').split(':');
  const p = parseRef(a), q = parseRef(b || a);
  if (!p || !q) return null;
  return { r1: Math.min(p.r, q.r), c1: Math.min(p.c, q.c), r2: Math.max(p.r, q.r), c2: Math.max(p.c, q.c) };
}

/* ───────────── 색 ───────────── */

// 엑셀 기본 색상표(indexed 0~63). 64는 글자색(검정), 65는 바탕색(흰색)
const INDEXED = ('000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF 000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF ' +
  '800000 008000 000080 808000 800080 008080 C0C0C0 808080 9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF ' +
  '000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF 00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 ' +
  '3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 003366 339966 003300 333300 993300 993366 333399 333333').split(' ');

function themeColors(xml) {
  // 엑셀은 테마 색 번호 0·1, 2·3의 순서를 바꿔 쓴다(0=lt1 바탕, 1=dk1 글자)
  const order = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
  const fallback = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];
  if (!xml) return fallback;
  const scheme = deep(parseXml(xml), 'clrScheme')[0];
  return order.map((name, i) => {
    const el = kid(scheme, name);
    const c = el && el.children[0];
    if (!c) return fallback[i];
    return (c.name === 'sysClr' ? attr(c, 'lastClr') : attr(c, 'val')) || fallback[i];
  });
}

function resolveColor(el, theme) {
  if (!el) return null;
  let hex = null;
  if (attr(el, 'rgb')) hex = attr(el, 'rgb').slice(-6);
  else if (attr(el, 'theme') != null) hex = theme[Number(attr(el, 'theme'))] || null;
  else if (attr(el, 'indexed') != null) {
    const i = Number(attr(el, 'indexed'));
    hex = i === 64 ? '000000' : i === 65 ? 'FFFFFF' : INDEXED[i] || null;
  } else if (attr(el, 'auto') === '1') hex = '000000';
  if (!hex) return null;
  let rgb = [0, 2, 4].map(i => parseInt(hex.slice(i, i + 2), 16));
  const tint = Number(attr(el, 'tint') || 0);
  if (tint) rgb = rgb.map(v => Math.round(tint > 0 ? v + (255 - v) * tint : v * (1 + tint)));
  return rgb;
}

/* ───────────── 스타일 ───────────── */

const BUILTIN_FMT = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00',
  14: 'yyyy-mm-dd', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM',
  20: 'h:mm', 21: 'h:mm:ss', 22: 'yyyy-mm-dd h:mm', 49: '@',
};
for (const id of [27, 28, 29, 30, 31, 34, 35, 36, 50, 51, 52, 53, 54, 57, 58]) BUILTIN_FMT[id] = 'yyyy"년" m"월" d"일"';
for (const id of [32, 33, 55, 56]) BUILTIN_FMT[id] = 'h"시" mm"분"';

function parseStyles(xml, theme) {
  const doc = parseXml(xml || '<styleSheet/>');
  const ss = kid(doc, 'styleSheet') || doc;
  const numFmts = { ...BUILTIN_FMT };
  for (const f of kids(kid(ss, 'numFmts'), 'numFmt')) numFmts[Number(attr(f, 'numFmtId'))] = attr(f, 'formatCode');
  const flag = (el, name) => {
    const k = kid(el, name);
    return !!k && attr(k, 'val') !== '0' && attr(k, 'val') !== 'false' && attr(k, 'val') !== 'none';
  };
  const fonts = kids(kid(ss, 'fonts'), 'font').map(f => ({
    name: attr(kid(f, 'name'), 'val') || '',
    sz: Number(attr(kid(f, 'sz'), 'val') || 11),
    b: flag(f, 'b'), i: flag(f, 'i'), u: flag(f, 'u'), strike: flag(f, 'strike'),
    color: resolveColor(kid(f, 'color'), theme),
  }));
  const fills = kids(kid(ss, 'fills'), 'fill').map(f => {
    const p = kid(f, 'patternFill');
    if (!p || attr(p, 'patternType') !== 'solid') return null;
    return resolveColor(kid(p, 'fgColor'), theme) || resolveColor(kid(p, 'bgColor'), theme);
  });
  const borders = kids(kid(ss, 'borders'), 'border').map(b => {
    const side = name => {
      const el = kid(b, name);
      const style = attr(el, 'style');
      if (!style || style === 'none') return null;
      return { style, color: resolveColor(kid(el, 'color'), theme) || [0, 0, 0] };
    };
    return { left: side('left') || side('start'), right: side('right') || side('end'), top: side('top'), bottom: side('bottom') };
  });
  const xfs = kids(kid(ss, 'cellXfs'), 'xf').map(x => {
    const al = kid(x, 'alignment');
    const fmtId = Number(attr(x, 'numFmtId') || 0);
    return {
      font: fonts[Number(attr(x, 'fontId') || 0)] || fonts[0] || { name: '', sz: 11 },
      fill: fills[Number(attr(x, 'fillId') || 0)] || null,
      border: borders[Number(attr(x, 'borderId') || 0)] || {},
      numFmt: numFmts[fmtId] || 'General',
      align: {
        h: attr(al, 'horizontal') || 'general',
        v: attr(al, 'vertical') || 'bottom',
        wrap: attr(al, 'wrapText') === '1' || attr(al, 'wrapText') === 'true',
        shrink: attr(al, 'shrinkToFit') === '1' || attr(al, 'shrinkToFit') === 'true',
        indent: Number(attr(al, 'indent') || 0),
      },
    };
  });
  // '표준' 스타일 글꼴: 열 너비 단위(가장 넓은 숫자 폭)를 정한다
  const normalXf = kids(kid(ss, 'cellStyleXfs'), 'xf')[0];
  const normalFont = fonts[Number(attr(normalXf, 'fontId') || 0)] || fonts[0] || { name: 'Calibri', sz: 11 };
  return { xfs, normalFont };
}

/* ───────────── 숫자·날짜 표시 ───────────── */

function serialToDate(n, date1904) {
  const ms = Math.round((n + (date1904 ? 1462 : 0) - 25569) * 86400000);
  return new Date(ms); // UTC 기준으로 읽는다
}

function formatDate(n, code, date1904) {
  const d = serialToDate(n, date1904);
  const p2 = v => String(v).padStart(2, '0');
  const parts = { yyyy: d.getUTCFullYear(), yy: p2(d.getUTCFullYear() % 100), mm: p2(d.getUTCMonth() + 1), m: d.getUTCMonth() + 1,
    dd: p2(d.getUTCDate()), d: d.getUTCDate(), hh: p2(d.getUTCHours()), h: d.getUTCHours(), ss: p2(d.getUTCSeconds()) };
  const days = ['일', '월', '화', '수', '목', '금', '토'];
  let out = '', lastWasHour = false;
  const re = /"([^"]*)"|\\(.)|(yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|aaaa|aaa|dd|d|hh|h|ss|s|AM\/PM)|(.)/gi;
  let t;
  while ((t = re.exec(code))) {
    if (t[1] != null) { out += t[1]; continue; }
    if (t[2] != null) { out += t[2]; continue; }
    if (t[4] != null) { out += t[4]; continue; }
    const tok = t[3].toLowerCase();
    if (tok === 'mm' || tok === 'm') {
      const rest = code.slice(re.lastIndex).replace(/"[^"]*"/g, '').replace(/^[^A-Za-z]*/, '');
      const minute = lastWasHour || /^s/i.test(rest);
      out += minute ? p2(d.getUTCMinutes()) : parts[tok];
    } else if (tok === 'aaa' || tok === 'ddd') out += days[d.getUTCDay()];
    else if (tok === 'aaaa' || tok === 'dddd') out += days[d.getUTCDay()] + '요일';
    else if (tok === 'mmm' || tok === 'mmmm') out += (d.getUTCMonth() + 1) + '월';
    else if (tok === 's') out += d.getUTCSeconds();
    else if (tok === 'am/pm') out += d.getUTCHours() < 12 ? '오전' : '오후';
    else out += parts[tok];
    lastWasHour = tok === 'hh' || tok === 'h';
  }
  return out;
}

function formatNumber(v, code, date1904) {
  const section = String(code || 'General').split(';')[0].replace(/\[[^\]]*\]/g, '');
  if (/^general$/i.test(section.trim()) || section === '@') {
    if (Number.isInteger(v)) return String(v);
    return String(Number(v.toPrecision(10)));
  }
  const bare = section.replace(/"[^"]*"/g, '').replace(/\\./g, '');
  if (/[ymdh]/i.test(bare) && !/[#0?]/.test(bare)) return formatDate(v, section, date1904);
  const percent = section.includes('%');
  const num = percent ? v * 100 : v;
  const dec = (/\.([0#]+)/.exec(section) || [, ''])[1].length;
  let s = Math.abs(num).toFixed(dec);
  if (section.includes(',')) {
    const [i, f] = s.split('.');
    s = i.replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (f ? '.' + f : '');
  }
  const lit = section.replace(/"([^"]*)"/g, '$1');
  const prefix = (/^[^#0.,%]*/.exec(lit) || [''])[0].replace(/\\/g, '');
  const suffix = (/[^#0.,%]*$/.exec(lit) || [''])[0].replace(/\\/g, '');
  return (num < 0 ? '-' : '') + prefix + s + (percent ? '%' : '') + suffix;
}

/* ───────────── 통합 문서 ───────────── */

function bad(message) { return Object.assign(new Error(message), { status: 400 }); }

function openWorkbook(buf) {
  const zip = readZip(buf);
  const text = p => (zip[p] ? zip[p]().toString('utf8') : null);
  const relsOf = file => {
    const rels = {};
    const xml = text(path.join(path.dirname(file), '_rels', path.basename(file) + '.rels'));
    if (!xml) return rels;
    for (const r of deep(parseXml(xml), 'Relationship')) {
      const target = attr(r, 'Target') || '';
      rels[attr(r, 'Id')] = attr(r, 'TargetMode') === 'External' ? null
        : target.startsWith('/') ? target.slice(1) : path.normalize(path.join(path.dirname(file), target));
    }
    return rels;
  };

  const wbXml = text('xl/workbook.xml');
  if (!wbXml) throw bad('엑셀(.xlsx) 파일이 아닙니다. 엑셀에서 "다른 이름으로 저장 → Excel 통합 문서(.xlsx)"로 저장해 주세요.');
  const wb = parseXml(wbXml);
  const date1904 = ['1', 'true'].includes(attr(deep(wb, 'workbookPr')[0], 'date1904'));
  const wbRels = relsOf('xl/workbook.xml');
  const sheets = deep(wb, 'sheet').map((s, i) => ({
    name: attr(s, 'name'), index: i, hidden: !!attr(s, 'state') && attr(s, 'state') !== 'visible',
    file: wbRels[attrLocal(s, 'id')],
  })).filter(s => s.file && zip[s.file]);
  if (!sheets.length) throw bad('엑셀 파일에서 시트를 찾지 못했습니다.');

  const printAreas = {};
  for (const dn of deep(wb, 'definedName')) {
    if (attr(dn, 'name') === '_xlnm.Print_Area' && attr(dn, 'localSheetId') != null) {
      const range = parseRange(dn.text.split(',')[0].replace(/'/g, ''));
      if (range) printAreas[Number(attr(dn, 'localSheetId'))] = range;
    }
  }

  const shared = [];
  const ssXml = text('xl/sharedStrings.xml');
  if (ssXml) {
    for (const si of deep(parseXml(ssXml), 'si')) {
      // 윗주(rPh) 안의 글자는 빼고 본문 글자만 모은다
      const collect = el => el.children.map(c => (c.name === 't' ? c.text : c.name === 'rPh' || c.name === 'phoneticPr' ? '' : collect(c))).join('');
      shared.push(unescapeOoxml(collect(si)));
    }
  }
  const theme = themeColors(text('xl/theme/theme1.xml'));
  const styles = parseStyles(text('xl/styles.xml'), theme);

  function readSheet(sheet) {
    const doc = parseXml(text(sheet.file));
    const ws = kid(doc, 'worksheet') || doc;
    const fmt = kid(ws, 'sheetFormatPr');
    const model = {
      name: sheet.name,
      defaultRowHeight: Number(attr(fmt, 'defaultRowHeight') || 15),
      defaultColWidth: attr(fmt, 'defaultColWidth') != null ? Number(attr(fmt, 'defaultColWidth')) : null,
      baseColWidth: Number(attr(fmt, 'baseColWidth') || 8),
      cols: kids(kid(ws, 'cols'), 'col').map(c => ({
        min: Number(attr(c, 'min')), max: Number(attr(c, 'max')), width: Number(attr(c, 'width') || 0),
        hidden: attr(c, 'hidden') === '1', style: Number(attr(c, 'style') || 0),
      })),
      rows: {}, cells: {}, merges: [], images: [], rowBreaks: [],
      dimension: null, printArea: printAreas[sheet.index] || null,
      page: {}, normalFont: styles.normalFont,
    };
    const dim = kid(ws, 'dimension');
    if (dim) model.dimension = parseRange(attr(dim, 'ref'));
    let maxR = 0, maxC = 0;
    for (const row of kids(kid(ws, 'sheetData'), 'row')) {
      const r = Number(attr(row, 'r'));
      model.rows[r] = {
        ht: attr(row, 'ht') != null ? Number(attr(row, 'ht')) : null,
        hidden: attr(row, 'hidden') === '1',
        s: attr(row, 'customFormat') === '1' ? Number(attr(row, 's') || 0) : null,
      };
      let nextC = 1;
      for (const c of kids(row, 'c')) {
        const ref = attr(c, 'r') ? parseRef(attr(c, 'r')) : { r, c: nextC };
        nextC = ref.c + 1;
        const t = attr(c, 't') || 'n';
        const s = Number(attr(c, 's') || 0);
        const v = kid(c, 'v');
        let value = null;
        if (t === 's' && v) value = shared[Number(v.text)] ?? '';
        else if (t === 'inlineStr') value = unescapeOoxml(deep(kid(c, 'is'), 't').map(x => x.text).join(''));
        else if (t === 'str' && v) value = unescapeOoxml(v.text);
        else if (t === 'b' && v) value = v.text === '1' ? 'TRUE' : 'FALSE';
        else if (t === 'e' && v) value = v.text;
        else if (v && v.text !== '') value = formatNumber(Number(v.text), (styles.xfs[s] || {}).numFmt, date1904);
        model.cells[`${ref.r},${ref.c}`] = { r: ref.r, c: ref.c, s, value, numeric: t === 'n' && value != null };
        if (ref.r > maxR) maxR = ref.r;
        if (ref.c > maxC) maxC = ref.c;
      }
    }
    if (!model.dimension) model.dimension = { r1: 1, c1: 1, r2: Math.max(1, maxR), c2: Math.max(1, maxC) };
    for (const m of kids(kid(ws, 'mergeCells'), 'mergeCell')) {
      const range = parseRange(attr(m, 'ref'));
      if (range) model.merges.push(range);
    }
    const margins = kid(ws, 'pageMargins');
    const setup = kid(ws, 'pageSetup');
    const setupPr = kid(kid(ws, 'sheetPr'), 'pageSetUpPr');
    const opts = kid(ws, 'printOptions');
    model.page = {
      left: Number(attr(margins, 'left') ?? 0.7), right: Number(attr(margins, 'right') ?? 0.7),
      top: Number(attr(margins, 'top') ?? 0.75), bottom: Number(attr(margins, 'bottom') ?? 0.75),
      paperSize: Number(attr(setup, 'paperSize') || 9),
      landscape: attr(setup, 'orientation') === 'landscape',
      scale: Number(attr(setup, 'scale') || 100),
      fitToPage: ['1', 'true'].includes(attr(setupPr, 'fitToPage')),
      fitToWidth: Number(attr(setup, 'fitToWidth') ?? 1),
      fitToHeight: Number(attr(setup, 'fitToHeight') ?? 1),
      hCentered: ['1', 'true'].includes(attr(opts, 'horizontalCentered')),
      vCentered: ['1', 'true'].includes(attr(opts, 'verticalCentered')),
    };
    for (const brk of kids(kid(ws, 'rowBreaks'), 'brk')) model.rowBreaks.push(Number(attr(brk, 'id')));

    // 그림
    const rels = relsOf(sheet.file);
    for (const d of kids(ws, 'drawing')) {
      const file = rels[attrLocal(d, 'id')];
      if (!file || !zip[file]) continue;
      const dRels = relsOf(file);
      const dr = parseXml(text(file));
      const anchors = [...deep(dr, 'twoCellAnchor'), ...deep(dr, 'oneCellAnchor'), ...deep(dr, 'absoluteAnchor')];
      for (const a of anchors) {
        for (const pic of deep(a, 'pic')) {
          const blip = deep(pic, 'blip')[0];
          const media = dRels[attrLocal(blip, 'embed')];
          if (!media || !zip[media]) continue;
          const data = zip[media]();
          const type = data[0] === 0x89 && data[1] === 0x50 ? 'png' : data[0] === 0xff && data[1] === 0xd8 ? 'jpeg' : null;
          if (!type) continue; // EMF 등 지원하지 않는 그림은 건너뛴다
          const pos = el => el && {
            col: Number(kid(el, 'col').text), colOff: Number(kid(el, 'colOff').text),
            row: Number(kid(el, 'row').text), rowOff: Number(kid(el, 'rowOff').text),
          };
          const src = deep(pic, 'srcRect')[0];
          const crop = k => (src && attr(src, k) ? Number(attr(src, k)) / 100000 : 0);
          const anchorExt = kid(a, 'ext');
          const xfrmExt = deep(kid(pic, 'spPr'), 'ext').find(e => attr(e, 'cx') != null);
          const ext = anchorExt || xfrmExt;
          const absPos = kid(a, 'pos');
          model.images.push({
            kind: a.name, type, data,
            from: pos(kid(a, 'from')), to: a.name === 'twoCellAnchor' ? pos(kid(a, 'to')) : null,
            pos: absPos ? { x: Number(attr(absPos, 'x')), y: Number(attr(absPos, 'y')) } : null,
            ext: ext ? { cx: Number(attr(ext, 'cx')), cy: Number(attr(ext, 'cy')) } : null,
            crop: { l: crop('l'), t: crop('t'), r: crop('r'), b: crop('b') },
          });
        }
      }
    }
    return model;
  }

  return { sheets, styles, readSheet: name => readSheet(sheets.find(s => s.name === name) || sheets[0]) };
}

module.exports = { openWorkbook, parseXml, parseRange, parseRef, colIndex, formatNumber };
