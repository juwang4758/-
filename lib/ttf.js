'use strict';
/**
 * TrueType 글꼴 읽기·부분 집합 만들기 (외부 패키지 없음)
 *
 * PDF에 글꼴 전체(3MB)를 넣지 않고, 문서에 실제로 쓴 글자만 남긴 글꼴을 넣는다.
 * 글리프 번호는 그대로 두고 쓰지 않은 글리프의 모양 데이터만 비운다(GID 유지 방식).
 * 그래서 PDF 쪽에서는 글리프 번호를 그대로 문자 코드(CID)로 쓸 수 있다.
 */

const KEEP_TABLES = ['OS/2', 'cvt ', 'fpgm', 'glyf', 'head', 'hhea', 'hmtx', 'loca', 'maxp', 'name', 'post', 'prep'];

function bad(message) { return new Error(message); }

class TrueTypeFont {
  constructor(buf) {
    if (buf.readUInt32BE(0) !== 0x00010000 && buf.toString('latin1', 0, 4) !== 'true') throw bad('TrueType 글꼴이 아닙니다.');
    this.buf = buf;
    this.tables = {};
    const n = buf.readUInt16BE(4);
    for (let i = 0; i < n; i++) {
      const o = 12 + i * 16;
      this.tables[buf.toString('latin1', o, o + 4)] = { off: buf.readUInt32BE(o + 8), len: buf.readUInt32BE(o + 12) };
    }
    for (const t of ['head', 'hhea', 'hmtx', 'maxp', 'loca', 'glyf']) if (!this.tables[t]) throw bad(`글꼴에 ${t} 표가 없습니다.`);
    const head = this.tables.head.off;
    this.unitsPerEm = buf.readUInt16BE(head + 18);
    this.bbox = [buf.readInt16BE(head + 36), buf.readInt16BE(head + 38), buf.readInt16BE(head + 40), buf.readInt16BE(head + 42)];
    this.longLoca = buf.readInt16BE(head + 50) === 1;
    this.numGlyphs = buf.readUInt16BE(this.tables.maxp.off + 4);
    const hhea = this.tables.hhea.off;
    this.ascent = buf.readInt16BE(hhea + 4);
    this.descent = buf.readInt16BE(hhea + 6);
    this.numHMetrics = buf.readUInt16BE(hhea + 34);
    const os2 = this.tables['OS/2'];
    this.capHeight = os2 && buf.readUInt16BE(os2.off) >= 2 ? buf.readInt16BE(os2.off + 88) : Math.round(this.ascent * 0.8);
    const post = this.tables.post;
    this.italicAngle = post ? buf.readInt32BE(post.off + 4) / 65536 : 0;
    this.postscriptName = this.readName(6) || 'Font';
    this.cmap = this.tables.cmap ? this.readCmap() : new Map(); // PDF용 부분 집합에는 문자표가 없다
    this.widthCache = new Map();
  }

  readName(id) {
    const t = this.tables.name;
    if (!t) return null;
    const b = this.buf, count = b.readUInt16BE(t.off + 2), strings = t.off + b.readUInt16BE(t.off + 4);
    for (let i = 0; i < count; i++) {
      const o = t.off + 6 + i * 12;
      const platform = b.readUInt16BE(o), nameId = b.readUInt16BE(o + 6), len = b.readUInt16BE(o + 8), off = b.readUInt16BE(o + 10);
      if (nameId !== id) continue;
      const raw = b.subarray(strings + off, strings + off + len);
      if (platform === 3 || platform === 0) {
        let s = '';
        for (let k = 0; k + 1 < raw.length; k += 2) s += String.fromCharCode(raw.readUInt16BE(k));
        return s.replace(/[^\x21-\x7e]/g, '');
      }
      if (platform === 1) return raw.toString('latin1').replace(/[^\x21-\x7e]/g, '');
    }
    return null;
  }

  /** 유니코드 → 글리프 번호 표. 형식 4(BMP)와 12(전체)를 읽는다 */
  readCmap() {
    const b = this.buf, base = this.tables.cmap.off, n = b.readUInt16BE(base + 2);
    let best = null;
    for (let i = 0; i < n; i++) {
      const platform = b.readUInt16BE(base + 4 + i * 8), enc = b.readUInt16BE(base + 6 + i * 8);
      const off = base + b.readUInt32BE(base + 8 + i * 8), format = b.readUInt16BE(off);
      const score = format === 12 ? 3 : format === 4 && (platform === 3 && enc === 1 || platform === 0) ? 2 : 0;
      if (score && (!best || score > best.score)) best = { off, format, score };
    }
    if (!best) throw bad('글꼴의 문자표(cmap)를 읽을 수 없습니다.');
    const map = new Map();
    const o = best.off;
    if (best.format === 4) {
      const segX2 = b.readUInt16BE(o + 6);
      const ends = o + 14, starts = ends + segX2 + 2, deltas = starts + segX2, ranges = deltas + segX2;
      for (let s = 0; s < segX2; s += 2) {
        const end = b.readUInt16BE(ends + s), start = b.readUInt16BE(starts + s);
        const delta = b.readInt16BE(deltas + s), rangeOff = b.readUInt16BE(ranges + s);
        for (let c = start; c <= end && c !== 0xffff; c++) {
          let g;
          if (rangeOff === 0) g = (c + delta) & 0xffff;
          else {
            const p = ranges + s + rangeOff + (c - start) * 2;
            g = b.readUInt16BE(p);
            if (g) g = (g + delta) & 0xffff;
          }
          if (g) map.set(c, g);
        }
      }
    } else {
      const groups = b.readUInt32BE(o + 12);
      for (let i = 0; i < groups; i++) {
        const p = o + 16 + i * 12;
        const start = b.readUInt32BE(p), end = b.readUInt32BE(p + 4), g0 = b.readUInt32BE(p + 8);
        for (let c = start; c <= end; c++) map.set(c, g0 + (c - start));
      }
    }
    return map;
  }

  glyphFor(codePoint) { return this.cmap.get(codePoint) || 0; }

  /** 글리프 폭(글꼴 단위) */
  advance(gid) {
    const b = this.buf, h = this.tables.hmtx.off;
    const i = Math.min(gid, this.numHMetrics - 1);
    return b.readUInt16BE(h + i * 4);
  }

  /** 글자 크기 1에서의 문자열 폭 */
  measure(text) {
    let w = 0;
    for (const ch of text) {
      let v = this.widthCache.get(ch);
      if (v === undefined) { v = this.advance(this.glyphFor(ch.codePointAt(0))) / this.unitsPerEm; this.widthCache.set(ch, v); }
      w += v;
    }
    return w;
  }

  glyphRange(gid) {
    const b = this.buf, l = this.tables.loca.off;
    if (this.longLoca) return [b.readUInt32BE(l + gid * 4), b.readUInt32BE(l + gid * 4 + 4)];
    return [b.readUInt16BE(l + gid * 2) * 2, b.readUInt16BE(l + gid * 2 + 2) * 2];
  }

  /** 복합 글리프가 가져다 쓰는 부품 글리프까지 포함한 집합 */
  closeOver(gids) {
    const out = new Set([0, ...gids]);
    const stack = [...out];
    const b = this.buf, glyf = this.tables.glyf.off;
    while (stack.length) {
      const g = stack.pop();
      if (g >= this.numGlyphs) continue;
      const [s, e] = this.glyphRange(g);
      if (e - s < 10 || b.readInt16BE(glyf + s) >= 0) continue;
      let p = glyf + s + 10;
      for (;;) {
        const flags = b.readUInt16BE(p), comp = b.readUInt16BE(p + 2);
        if (!out.has(comp)) { out.add(comp); stack.push(comp); }
        p += 4 + (flags & 0x0001 ? 4 : 2);
        if (flags & 0x0008) p += 2; else if (flags & 0x0040) p += 4; else if (flags & 0x0080) p += 8;
        if (!(flags & 0x0020)) break;
      }
    }
    return out;
  }

  /** 쓴 글리프만 남긴 글꼴 파일 */
  subset(gids) {
    const keep = this.closeOver(gids);
    const b = this.buf, glyfOff = this.tables.glyf.off;
    const parts = [];
    const loca = Buffer.alloc((this.numGlyphs + 1) * 4);
    let pos = 0;
    for (let g = 0; g < this.numGlyphs; g++) {
      loca.writeUInt32BE(pos, g * 4);
      if (!keep.has(g)) continue;
      const [s, e] = this.glyphRange(g);
      if (e <= s) continue;
      const data = b.subarray(glyfOff + s, glyfOff + e);
      parts.push(data);
      pos += data.length;
      const pad = (4 - (data.length % 4)) % 4;
      if (pad) { parts.push(Buffer.alloc(pad)); pos += pad; }
    }
    loca.writeUInt32BE(pos, this.numGlyphs * 4);
    const out = {};
    for (const tag of KEEP_TABLES) {
      const t = this.tables[tag];
      if (t) out[tag] = Buffer.from(b.subarray(t.off, t.off + t.len));
    }
    out.glyf = Buffer.concat(parts);
    out.loca = loca;
    out.head.writeInt16BE(1, 50); // loca를 4바이트 형식으로 쓴다
    out.head.writeUInt32BE(0, 8);
    if (out.post && out.post.length >= 32) { // 글리프 이름은 필요 없으므로 버전 3으로 줄인다
      out.post = out.post.subarray(0, 32);
      out.post.writeUInt32BE(0x00030000, 0);
    }
    return buildSfnt(out);
  }
}

function checksum(buf) {
  let sum = 0;
  const padded = buf.length % 4 ? Buffer.concat([buf, Buffer.alloc(4 - (buf.length % 4))]) : buf;
  for (let i = 0; i < padded.length; i += 4) sum = (sum + padded.readUInt32BE(i)) >>> 0;
  return sum;
}

function buildSfnt(tables) {
  const tags = Object.keys(tables).sort();
  const n = tags.length;
  let pow = 1, log = 0;
  while (pow * 2 <= n) { pow *= 2; log++; }
  const header = Buffer.alloc(12 + n * 16);
  header.writeUInt32BE(0x00010000, 0);
  header.writeUInt16BE(n, 4);
  header.writeUInt16BE(pow * 16, 6);
  header.writeUInt16BE(log, 8);
  header.writeUInt16BE(n * 16 - pow * 16, 10);
  const bodies = [];
  let offset = header.length;
  tags.forEach((tag, i) => {
    const data = tables[tag];
    const o = 12 + i * 16;
    header.write(tag, o, 4, 'latin1');
    header.writeUInt32BE(checksum(data), o + 4);
    header.writeUInt32BE(offset, o + 8);
    header.writeUInt32BE(data.length, o + 12);
    const pad = (4 - (data.length % 4)) % 4;
    bodies.push(data, Buffer.alloc(pad));
    offset += data.length + pad;
  });
  const font = Buffer.concat([header, ...bodies]);
  const headOff = header.readUInt32BE(12 + tags.indexOf('head') * 16 + 8);
  font.writeUInt32BE((0xb1b0afba - checksum(font)) >>> 0, headOff + 8);
  return font;
}

module.exports = { TrueTypeFont };
