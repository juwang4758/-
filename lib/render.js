'use strict';
/**
 * 엑셀 양식 시트 → PDF
 *
 * 엑셀이 인쇄하는 방식을 따라 그린다.
 *  - 인쇄 영역(없으면 내용이 있는 범위)을 용지 폭에 맞춰 줄인다("한 페이지에 맞추기"와 같은 배율 계산).
 *  - 병합한 칸, 가로·세로 정렬, 자동 줄바꿈, 테두리, 채우기, 그림(자르기 포함)을 반영한다.
 *  - 칸 안의 [교회명]·[문서번호]·[날짜]를 교회마다 바꿔 넣는다.
 * 글꼴은 양식에 지정된 글꼴과 상관없이 나눔명조(굵게 포함)로 그린다.
 */
const { PdfDocument } = require('./pdf');

const EMU_PER_PT = 12700;
const LINE_HEIGHT = 1.3; // 줄 간격(글자 크기 배수)
const PAPER = { 1: [612, 792], 5: [612, 1008], 8: [841.89, 1190.55], 9: [595.28, 841.89], 11: [419.53, 595.28], 13: [515.91, 728.5], 70: [297.64, 419.53] };
const PLACEHOLDER = /\[\s*(교회명|문서번호|날짜|시행일자)\s*\]/g;

function fillPlaceholders(text, vars) {
  return text.replace(PLACEHOLDER, (all, key) => {
    const v = vars[key === '시행일자' ? '날짜' : key];
    return v == null || v === '' ? all : String(v);
  });
}

/* ───────────── 글꼴에 없는 글자 ───────────── */

const ROMAN = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII'];
const UNITS = { '㎡': 'm²', '㎥': 'm³', '㎏': 'kg', '㎎': 'mg', '㎞': 'km', '㎝': 'cm', '㎜': 'mm', '㎖': 'mL', '㎗': 'dL', 'ℓ': 'L', '℃': '°C', '℉': '°F' };
const SIMILAR = { '∼': '~', '～': '~', '―': '—', '‧': '·', '✓': '√', '✔': '√', ' ': ' ', '　': ' ', '\t': '  ' };

function substitute(ch) {
  const cp = ch.codePointAt(0);
  if (cp >= 0x2460 && cp <= 0x2473) return `(${cp - 0x245f})`; // ①
  if (cp >= 0x2474 && cp <= 0x2487) return `(${cp - 0x2473})`; // ⑴
  if (cp >= 0x2488 && cp <= 0x249b) return `${cp - 0x2487}.`; // ⒈
  if (cp >= 0x2776 && cp <= 0x277f) return `(${cp - 0x2775})`; // ❶
  if (cp >= 0x2780 && cp <= 0x2789) return `(${cp - 0x277f})`; // ➀
  if (cp >= 0x278a && cp <= 0x2793) return `(${cp - 0x2789})`; // ➊
  if (cp >= 0x2160 && cp <= 0x216b) return ROMAN[cp - 0x2160];
  if (cp >= 0x2170 && cp <= 0x217b) return ROMAN[cp - 0x2170].toLowerCase();
  return UNITS[ch] || SIMILAR[ch] || null;
}

/** 글꼴에 없는 글자를 비슷한 글자로 바꾸고, 바꿀 수 없는 글자는 missing에 모은다 */
function fontSafe(text, ttf, missing) {
  let out = '';
  for (const ch of text.normalize('NFC')) {
    if (ch === '\n' || ttf.glyphFor(ch.codePointAt(0))) { out += ch; continue; }
    const sub = substitute(ch);
    const nfkc = ch.normalize('NFKC');
    const ok = s => s && [...s].every(c => ttf.glyphFor(c.codePointAt(0)));
    if (ok(sub)) out += sub;
    else if (nfkc !== ch && ok(nfkc)) out += nfkc;
    else if (/\s/.test(ch)) out += ' ';
    else if (/[\u0000-\u001f​-‏﻿]/.test(ch)) continue; // 보이지 않는 제어 문자
    else { out += ch; missing.add(ch); }
  }
  return out;
}

/* ───────────── 줄바꿈 ───────────── */

function wrapLines(text, measure, maxW) {
  const lines = [];
  for (const para of text.split('\n')) {
    let line = '';
    for (const token of para.split(/(\s+)/).filter(t => t !== '')) {
      if (measure(line + token) <= maxW) { line += token; continue; }
      if (/^\s+$/.test(token)) { lines.push(line); line = ''; continue; }
      if (line.trim()) { lines.push(line.replace(/\s+$/, '')); line = ''; }
      if (measure(token) <= maxW) { line = token; continue; }
      for (const ch of token) { // 칸보다 긴 낱말은 글자 단위로 나눈다
        if (line && measure(line + ch) > maxW) { lines.push(line); line = ''; }
        line += ch;
      }
    }
    lines.push(line);
  }
  return lines;
}

/* ───────────── 시트 배치 계산 ───────────── */

function maxDigitWidthPx(font) {
  const px = ((font && font.sz) || 11) * 96 / 72;
  const latin = /calibri|arial|helvetica|verdana|tahoma|segoe|times|cambria/i.test((font && font.name) || '');
  return Math.max(5, Math.round(px * (latin ? 0.47 : 0.55)));
}

function layoutSheet(sheet, styles) {
  const mdw = maxDigitWidthPx(sheet.normalFont);
  const charsToPt = w => Math.trunc(((256 * w + Math.trunc(128 / mdw)) / 256) * mdw) * 0.75;
  const defaultCol = sheet.defaultColWidth != null ? charsToPt(sheet.defaultColWidth) : Math.trunc(sheet.baseColWidth * mdw + 5) * 0.75;
  const colW = c => {
    const def = sheet.cols.find(d => c >= d.min && c <= d.max);
    if (!def) return defaultCol;
    return def.hidden ? 0 : charsToPt(def.width);
  };
  const rowH = r => {
    const row = sheet.rows[r];
    if (row && row.hidden) return 0;
    return row && row.ht != null ? row.ht : sheet.defaultRowHeight;
  };
  const xfOf = (r, c) => {
    const cell = sheet.cells[`${r},${c}`];
    if (cell) return styles.xfs[cell.s] || styles.xfs[0];
    const row = sheet.rows[r];
    if (row && row.s != null) return styles.xfs[row.s] || styles.xfs[0];
    const def = sheet.cols.find(d => c >= d.min && c <= d.max);
    return styles.xfs[def ? def.style : 0] || styles.xfs[0];
  };
  const hasBorder = xf => xf && xf.border && (xf.border.left || xf.border.right || xf.border.top || xf.border.bottom);

  // 인쇄할 범위: 인쇄 영역이 있으면 그것, 없으면 값·테두리·채우기·그림이 있는 범위
  let range = sheet.printArea;
  if (!range) {
    let r1 = Infinity, c1 = Infinity, r2 = 0, c2 = 0;
    const take = (r, c) => { r1 = Math.min(r1, r); c1 = Math.min(c1, c); r2 = Math.max(r2, r); c2 = Math.max(c2, c); };
    for (const cell of Object.values(sheet.cells)) {
      const xf = styles.xfs[cell.s] || styles.xfs[0];
      if ((cell.value != null && cell.value !== '') || hasBorder(xf) || (xf && xf.fill)) take(cell.r, cell.c);
    }
    for (const m of sheet.merges) {
      const inside = Object.values(sheet.cells).some(cell => cell.r >= m.r1 && cell.r <= m.r2 && cell.c >= m.c1 && cell.c <= m.c2 && cell.value != null && cell.value !== '');
      if (inside) { take(m.r1, m.c1); take(m.r2, m.c2); }
    }
    for (const img of sheet.images) if (img.from) take(img.from.row + 1, img.from.col + 1);
    range = r2 ? { r1, c1, r2, c2 } : { r1: 1, c1: 1, r2: 1, c2: 1 };
  }

  // 시트 원래 크기(pt)에서의 열·행 시작 위치
  const colX = [], rowY = [];
  let x = 0;
  const maxCol = Math.max(range.c2, ...sheet.images.map(i => (i.to ? i.to.col + 1 : i.from ? i.from.col + 1 : 1)));
  for (let c = 1; c <= maxCol + 1; c++) { colX[c] = x; x += colW(c); }
  let y = 0;
  const maxRow = Math.max(range.r2, ...sheet.images.map(i => (i.to ? i.to.row + 1 : i.from ? i.from.row + 1 : 1)));
  for (let r = 1; r <= maxRow + 1; r++) { rowY[r] = y; y += rowH(r); }

  const [pw, ph] = PAPER[sheet.page.paperSize] || PAPER[9];
  const pageW = sheet.page.landscape ? ph : pw, pageH = sheet.page.landscape ? pw : ph;
  const ml = sheet.page.left * 72, mr = sheet.page.right * 72, mt = sheet.page.top * 72, mb = sheet.page.bottom * 72;
  const printW = pageW - ml - mr, printH = pageH - mt - mb;
  const totalW = colX[range.c2 + 1] - colX[range.c1];
  const totalH = rowY[range.r2 + 1] - rowY[range.r1];
  let scale;
  if (sheet.page.fitToPage) {
    const fw = sheet.page.fitToWidth, fh = sheet.page.fitToHeight;
    scale = Math.min(1, fw > 0 ? (printW * fw) / totalW : 1, fh > 0 ? (printH * fh) / totalH : 1);
    scale = Math.max(0.1, Math.floor(scale * 100) / 100); // 엑셀처럼 % 단위로 내림
  } else scale = Math.min(4, Math.max(0.1, (sheet.page.scale || 100) / 100));

  // 쪽 나누기: 행이 인쇄 높이를 넘으면 다음 쪽으로
  const pages = [];
  let start = range.r1, used = 0;
  for (let r = range.r1; r <= range.r2; r++) {
    const h = rowH(r) * scale;
    if (r > start && (used + h > printH + 0.01 || sheet.rowBreaks.includes(r - 1))) { pages.push([start, r - 1]); start = r; used = 0; }
    used += h;
  }
  pages.push([start, range.r2]);

  const offsetX = sheet.page.hCentered ? Math.max(0, (printW - totalW * scale) / 2) : 0;
  return { range, colX, rowY, scale, pageW, pageH, ml, mt, printH, offsetX, pages, xfOf, mdw };
}

/* ───────────── 테두리 ───────────── */

const BORDER = { // 굵기(px), 우선순위
  hair: [0.5, 1], dotted: [1, 2], dashDotDot: [1, 3], dashDot: [1, 3], dashed: [1, 3], thin: [1, 4],
  mediumDashDotDot: [2, 5], slantDashDot: [2, 5], mediumDashDot: [2, 5], mediumDashed: [2, 5], medium: [2, 6],
  double: [3, 7], thick: [3, 8],
};

function collectEdges(sheet, L) {
  const { range, xfOf } = L;
  const edges = new Map();
  const interiorH = (r, c) => sheet.merges.some(m => c >= m.c1 && c <= m.c2 && r > m.r1 && r <= m.r2);
  const interiorV = (r, c) => sheet.merges.some(m => r >= m.r1 && r <= m.r2 && c > m.c1 && c <= m.c2);
  const put = (key, side) => {
    if (!side || !BORDER[side.style]) return;
    const old = edges.get(key);
    if (!old || BORDER[side.style][1] > BORDER[old.style][1]) edges.set(key, side);
  };
  for (let r = range.r1; r <= range.r2; r++) {
    for (let c = range.c1; c <= range.c2; c++) {
      const b = (xfOf(r, c) || {}).border || {};
      if (!interiorH(r, c)) put(`h,${r},${c}`, b.top);
      if (!interiorH(r + 1, c)) put(`h,${r + 1},${c}`, b.bottom);
      if (!interiorV(r, c)) put(`v,${r},${c}`, b.left);
      if (!interiorV(r, c + 1)) put(`v,${r},${c + 1}`, b.right);
    }
  }
  return edges;
}

function drawEdges(page, edges, L, rowFrom, rowTo, toX, toY) {
  const s = L.scale;
  const segs = [];
  for (const [key, side] of edges) {
    const [kind, rs, cs] = key.split(',');
    const r = Number(rs), c = Number(cs);
    if (kind === 'h' && (r < rowFrom || r > rowTo + 1)) continue;
    if (kind === 'v' && (r < rowFrom || r > rowTo)) continue;
    if (kind === 'h') segs.push({ kind, at: r, from: c, to: c + 1, side });
    else segs.push({ kind, at: c, from: r, to: r + 1, side });
  }
  // 같은 줄에 이어진 같은 모양의 선은 한 번에 긋는다
  segs.sort((a, b) => (a.kind < b.kind ? -1 : a.kind > b.kind ? 1 : a.at - b.at || a.from - b.from));
  const merged = [];
  for (const g of segs) {
    const last = merged[merged.length - 1];
    if (last && last.kind === g.kind && last.at === g.at && last.to === g.from && last.side.style === g.side.style && String(last.side.color) === String(g.side.color)) last.to = g.to;
    else merged.push({ ...g });
  }
  for (const g of merged) {
    const color = g.side.color || [0, 0, 0];
    if (color.every(v => v >= 250)) continue; // 흰 선은 인쇄해도 보이지 않는다
    const [px] = BORDER[g.side.style];
    const w = Math.max(0.25, px * 0.75 * s);
    const dashUnit = Math.max(0.5, 0.75 * s);
    const dash = /dashDotDot/i.test(g.side.style) ? [4, 2, 1, 2, 1, 2] : /dashDot/i.test(g.side.style) ? [4, 2, 1, 2]
      : /dashed/i.test(g.side.style) ? [4, 2] : g.side.style === 'dotted' || g.side.style === 'hair' ? [1, 1] : null;
    const pattern = dash && dash.map(v => v * dashUnit);
    const line = (off, width) => {
      if (g.kind === 'h') page.line(toX(g.from), toY(g.at) + off, toX(g.to), toY(g.at) + off, width, color, pattern);
      else page.line(toX(g.at) + off, toY(g.from), toX(g.at) + off, toY(g.to), width, color, pattern);
    };
    if (g.side.style === 'double') { const thin = Math.max(0.25, 0.75 * s); line(-thin, thin); line(thin, thin); }
    else line(0, w);
  }
}

/* ───────────── 그리기 ───────────── */

/**
 * sheet: xlsx.readSheet 결과, styles: 통합 문서 스타일, fonts: { regular, bold } TrueTypeFont
 * letters: [{ vars: { 교회명, 문서번호, 날짜 } }] — 교회 하나가 한 부(여러 쪽일 수 있음)
 * 돌려주는 값: { pdf: Buffer, missing: [글꼴에 없는 글자] }
 */
function renderPdf(sheet, styles, fonts, letters, opts = {}) {
  const L = layoutSheet(sheet, styles);
  const doc = new PdfDocument();
  if (opts.title) doc.info.title = opts.title;
  const pdfFonts = { regular: doc.addFont(fonts.regular), bold: doc.addFont(fonts.bold) };
  const missing = new Set();
  const edges = collectEdges(sheet, L);
  const s = L.scale;
  const anchorOf = (r, c) => sheet.merges.find(m => r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) || { r1: r, c1: c, r2: r, c2: c };
  const cells = Object.values(sheet.cells).filter(cell => {
    const m = anchorOf(cell.r, cell.c);
    return m.r1 === cell.r && m.c1 === cell.c && cell.r >= L.range.r1 && cell.r <= L.range.r2 && cell.c >= L.range.c1 && cell.c <= L.range.c2;
  });

  for (const letter of letters) {
    for (const [rowFrom, rowTo] of L.pages) {
      const page = doc.addPage(L.pageW, L.pageH);
      const toX = c => L.ml + L.offsetX + (L.colX[c] - L.colX[L.range.c1]) * s;
      const toY = r => L.mt + (L.rowY[r] - L.rowY[rowFrom]) * s;
      const ptX = x => L.ml + L.offsetX + (x - L.colX[L.range.c1]) * s;
      const ptY = y => L.mt + (y - L.rowY[rowFrom]) * s;

      // 채우기
      for (let r = rowFrom; r <= rowTo; r++) {
        for (let c = L.range.c1; c <= L.range.c2; c++) {
          const m = anchorOf(r, c);
          if (m.r1 !== r || m.c1 !== c) continue;
          const xf = L.xfOf(r, c);
          if (!xf || !xf.fill || xf.fill.every(v => v >= 250)) continue;
          page.fillRect(toX(m.c1), toY(m.r1), toX(Math.min(m.c2, L.range.c2) + 1) - toX(m.c1), toY(Math.min(m.r2, rowTo) + 1) - toY(m.r1), xf.fill);
        }
      }

      // 글자
      for (const cell of cells) {
        if (cell.r < rowFrom || cell.r > rowTo || cell.value == null || cell.value === '') continue;
        const m = anchorOf(cell.r, cell.c);
        const xf = styles.xfs[cell.s] || styles.xfs[0];
        const font = xf.font || {};
        const ttf = font.b ? fonts.bold : fonts.regular;
        const pf = font.b ? pdfFonts.bold : pdfFonts.regular;
        const text = fontSafe(fillPlaceholders(String(cell.value).replace(/\r\n?/g, '\n'), letter.vars || {}), ttf, missing);
        const bx = toX(m.c1), by = toY(m.r1);
        const bw = toX(Math.min(m.c2, L.range.c2) + 1) - bx, bh = toY(Math.min(m.r2, rowTo) + 1) - by;
        if (bw <= 0 || bh <= 0) continue; // 숨긴 행·열
        const padX = 2 * 0.75 * s, padY = 1 * s;
        const indent = (xf.align.indent || 0) * 3 * L.mdw * 0.75 * s;
        const availW = Math.max(1, bw - 2 * padX - indent);
        let size = (font.sz || 11) * s;
        const merged = m.r1 !== m.r2 || m.c1 !== m.c2;
        const measure = (str, sz) => ttf.measure(str) * sz;
        let lines;
        const layout = sz => (xf.align.wrap ? wrapLines(text, str => measure(str, sz), availW) : text.split('\n'));
        lines = layout(size);
        // 칸을 넘치면 글자를 줄여 잘리지 않게 한다(병합 칸·줄바꿈 칸·축소 맞춤)
        for (let k = 0; k < 40; k++) {
          const tooWide = !xf.align.wrap && (merged || xf.align.shrink) && lines.some(l => measure(l, size) > availW);
          const tooTall = (xf.align.wrap || merged) && lines.length * size * LINE_HEIGHT > bh - 2 * padY + 0.01;
          if ((!tooWide && !tooTall) || size <= (font.sz || 11) * s * 0.5) break;
          size *= 0.96;
          lines = layout(size);
        }
        const lineH = size * LINE_HEIGHT;
        const blockH = lines.length * lineH;
        const v = xf.align.v;
        let top = v === 'top' ? by + padY : v === 'center' || v === 'distributed' || v === 'justify' ? by + (bh - blockH) / 2 : by + bh - padY - blockH;
        const ascent = (ttf.ascent / ttf.unitsPerEm) * size;
        const emTop = (lineH - size * ((ttf.ascent - ttf.descent) / ttf.unitsPerEm)) / 2;
        let h = xf.align.h;
        if (h === 'general') h = cell.numeric ? 'right' : 'left';
        const color = font.color || [0, 0, 0]; // 흰 글씨로 숨겨 둔 칸은 PDF에서도 보이지 않게 둔다
        for (const line of lines) {
          const baseline = top + emTop + ascent;
          const lw = measure(line, size);
          let x = bx + padX + indent, charSpace = 0;
          if (h === 'center' || h === 'centerContinuous') x = bx + (bw - lw) / 2;
          else if (h === 'right') x = bx + bw - padX - lw;
          else if (h === 'distributed' && [...line].length > 1 && lw < availW) {
            charSpace = (availW - lw) / ([...line].length - 1);
          }
          page.text(pf, size, x, baseline, line, { color, charSpace });
          if (font.u || font.strike) {
            const total = lw + charSpace * Math.max(0, [...line].length - 1);
            const uy = font.u ? baseline + size * 0.12 : baseline - size * 0.3;
            page.line(x, uy, x + total, uy, Math.max(0.3, size * 0.05), color);
          }
          top += lineH;
        }
      }

      drawEdges(page, edges, L, rowFrom, rowTo, toX, toY);

      // 그림
      for (const img of sheet.images) {
        let x1, y1, x2, y2;
        if (img.pos) { x1 = img.pos.x / EMU_PER_PT; y1 = img.pos.y / EMU_PER_PT; }
        else if (img.from) { x1 = L.colX[img.from.col + 1] + img.from.colOff / EMU_PER_PT; y1 = L.rowY[img.from.row + 1] + img.from.rowOff / EMU_PER_PT; }
        else continue;
        if (img.to) { x2 = L.colX[img.to.col + 1] + img.to.colOff / EMU_PER_PT; y2 = L.rowY[img.to.row + 1] + img.to.rowOff / EMU_PER_PT; }
        else if (img.ext) { x2 = x1 + img.ext.cx / EMU_PER_PT; y2 = y1 + img.ext.cy / EMU_PER_PT; }
        else continue;
        const anchorRow = img.from ? img.from.row + 1 : L.range.r1;
        if (anchorRow < rowFrom || anchorRow > rowTo) continue;
        const box = { x: ptX(x1), y: ptY(y1), w: (x2 - x1) * s, h: (y2 - y1) * s };
        if (box.w <= 0 || box.h <= 0) continue;
        const { l, t, r, b } = img.crop;
        const fullW = box.w / Math.max(0.01, 1 - l - r), fullH = box.h / Math.max(0.01, 1 - t - b);
        const handle = doc.addImage(img.data, img.type);
        const cropped = l || t || r || b;
        page.image(handle, box.x - l * fullW, box.y - t * fullH, fullW, fullH, cropped ? box : null);
      }
    }
  }
  return { pdf: doc.toBuffer(), missing: [...missing], pages: L.pages.length, scale: s };
}

/** 양식에 들어 있는 자리 표시([교회명] 등)와 각 개수 */
function findPlaceholders(sheet) {
  const found = {};
  for (const cell of Object.values(sheet.cells)) {
    if (typeof cell.value !== 'string') continue;
    for (const m of cell.value.matchAll(PLACEHOLDER)) {
      const key = m[1] === '시행일자' ? '날짜' : m[1];
      found[key] = (found[key] || 0) + 1;
    }
  }
  return found;
}

/** "제목" 칸 오른쪽에 적힌 제목 */
function findLabeled(sheet, label) {
  const cells = Object.values(sheet.cells);
  const lab = cells.find(c => typeof c.value === 'string' && c.value.replace(/\s/g, '') === label);
  if (!lab) return '';
  const m = sheet.merges.find(g => lab.r >= g.r1 && lab.r <= g.r2 && lab.c >= g.c1 && lab.c <= g.c2);
  const after = m ? m.c2 : lab.c;
  const v = cells.filter(c => c.r === lab.r && c.c > after && c.value != null && String(c.value).trim()).sort((a, b) => a.c - b.c)[0];
  return v ? String(v.value).trim() : '';
}

module.exports = { renderPdf, findPlaceholders, findLabeled, fillPlaceholders, wrapLines, fontSafe, layoutSheet };
