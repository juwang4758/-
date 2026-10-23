'use strict';
/**
 * 작은 PDF 작성기 (외부 패키지 없음)
 *
 * 공문 한 장을 그리는 데 필요한 것만 있다: 한글 글꼴(Type0/CIDFontType2, 쓴 글자만 넣음),
 * 글자, 선, 색 채우기, 그림(PNG·JPEG). 좌표는 왼쪽 위가 (0, 0)인 pt 단위로 받는다.
 */
const zlib = require('zlib');
const crypto = require('crypto');
const { decodePng, jpegInfo } = require('./image');

const num = v => {
  const s = (Math.round(v * 1000) / 1000).toString();
  return s === '-0' ? '0' : s;
};
const hex4 = n => n.toString(16).padStart(4, '0');
function pdfString(s) { // PDF 정보 사전용 UTF-16BE 문자열
  const b = Buffer.from('﻿' + s, 'utf16le').swap16();
  return '<' + b.toString('hex') + '>';
}
const colorOps = (rgb, stroke) => rgb.map(v => num(v / 255)).join(' ') + (stroke ? ' RG' : ' rg');
// 같은 로고를 PDF마다 다시 풀고 압축하지 않도록, 압축한 그림 데이터를 원본 버퍼별로 기억한다
const encodedPng = new WeakMap();

class Page {
  constructor(doc, width, height) {
    this.doc = doc; this.width = width; this.height = height;
    this.ops = []; this.fonts = new Set(); this.images = new Set();
  }

  /** 글자 한 줄. y는 기준선(baseline) 위치 */
  text(font, size, x, y, str, opts = {}) {
    if (!str) return;
    this.fonts.add(font);
    const color = opts.color || [0, 0, 0];
    this.ops.push(`BT ${colorOps(color)} /${font.name} ${num(size)} Tf ${opts.charSpace ? num(opts.charSpace) + ' Tc ' : ''}1 0 0 1 ${num(x)} ${num(this.height - y)} Tm <${font.encode(str)}> Tj ET`);
  }

  line(x1, y1, x2, y2, width, color = [0, 0, 0], dash = null) {
    this.ops.push(`q ${colorOps(color, true)} ${num(width)} w 0 J ${dash ? `[${dash.map(num).join(' ')}] 0 d ` : ''}${num(x1)} ${num(this.height - y1)} m ${num(x2)} ${num(this.height - y2)} l S Q`);
  }

  fillRect(x, y, w, h, color) {
    this.ops.push(`q ${colorOps(color)} ${num(x)} ${num(this.height - y - h)} ${num(w)} ${num(h)} re f Q`);
  }

  /** 그림을 (x, y, w, h)에 그린다. clip이 있으면 그 사각형 밖은 잘라 낸다 */
  image(img, x, y, w, h, clip = null) {
    this.images.add(img);
    const c = clip ? `${num(clip.x)} ${num(this.height - clip.y - clip.h)} ${num(clip.w)} ${num(clip.h)} re W n ` : '';
    this.ops.push(`q ${c}${num(w)} 0 0 ${num(h)} ${num(x)} ${num(this.height - y - h)} cm /${img.name} Do Q`);
  }
}

class Font {
  constructor(doc, ttf, name) {
    this.doc = doc; this.ttf = ttf; this.name = name;
    this.used = new Map(); // 글리프 번호 → 글자(복사·검색용)
    this.ids = { type0: doc.alloc(), cid: doc.alloc(), desc: doc.alloc(), file: doc.alloc(), cmap: doc.alloc() };
  }

  encode(str) {
    let out = '';
    for (const ch of str) {
      const cp = ch.codePointAt(0);
      let gid = this.ttf.glyphFor(cp);
      if (!gid && (cp === 0xa0 || cp === 0x3000)) gid = this.ttf.glyphFor(0x20);
      if (gid && !this.used.has(gid)) this.used.set(gid, ch);
      out += hex4(gid);
    }
    return out;
  }

  finish() {
    const doc = this.doc, ttf = this.ttf;
    const gids = [...this.used.keys()].sort((a, b) => a - b);
    const tag = crypto.createHash('md5').update(gids.join(',') + ttf.postscriptName).digest()
      .subarray(0, 6).reduce((s, b) => s + String.fromCharCode(65 + (b % 26)), '');
    const base = `${tag}+${ttf.postscriptName}`;
    const k = 1000 / ttf.unitsPerEm;
    const widths = gids.map(g => `${g} [${Math.round(ttf.advance(g) * k)}]`).join(' ');
    const fontFile = ttf.subset(gids);
    doc.setStream(this.ids.file, `/Length1 ${fontFile.length}`, fontFile);
    doc.set(this.ids.desc, `<< /Type /FontDescriptor /FontName /${base} /Flags 6 /FontBBox [${ttf.bbox.map(v => Math.round(v * k)).join(' ')}] ` +
      `/ItalicAngle ${num(ttf.italicAngle)} /Ascent ${Math.round(ttf.ascent * k)} /Descent ${Math.round(ttf.descent * k)} ` +
      `/CapHeight ${Math.round(ttf.capHeight * k)} /StemV 80 /FontFile2 ${this.ids.file} 0 R >>`);
    doc.set(this.ids.cid, `<< /Type /Font /Subtype /CIDFontType2 /BaseFont /${base} ` +
      `/CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor ${this.ids.desc} 0 R ` +
      `/DW 1000 /W [${widths}] /CIDToGIDMap /Identity >>`);
    // 글자 복사·검색이 되도록 글리프 → 유니코드 표를 붙인다
    const entries = gids.map(g => {
      const u = Buffer.from(this.used.get(g), 'utf16le').swap16().toString('hex');
      return `<${hex4(g)}> <${u}>`;
    });
    let body = '';
    for (let i = 0; i < entries.length; i += 100) {
      const chunk = entries.slice(i, i + 100);
      body += `${chunk.length} beginbfchar\n${chunk.join('\n')}\nendbfchar\n`;
    }
    const cmap = '/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n' +
      '/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n' +
      '1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n' + body +
      'endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n';
    doc.setStream(this.ids.cmap, '', Buffer.from(cmap, 'latin1'));
    doc.set(this.ids.type0, `<< /Type /Font /Subtype /Type0 /BaseFont /${base} /Encoding /Identity-H ` +
      `/DescendantFonts [${this.ids.cid} 0 R] /ToUnicode ${this.ids.cmap} 0 R >>`);
  }
}

class PdfDocument {
  constructor() {
    this.objects = [null];
    this.pages = [];
    this.fontList = [];
    this.imageCache = new Map();
    this.info = {};
  }

  alloc() { this.objects.push(null); return this.objects.length - 1; }
  set(id, body) { this.objects[id] = Buffer.from(body, 'latin1'); }
  setStream(id, extraDict, data, compressed = null) {
    const body = compressed || zlib.deflateSync(data);
    this.objects[id] = Buffer.concat([
      Buffer.from(`<< /Length ${body.length} /Filter /FlateDecode ${extraDict} >>\nstream\n`, 'latin1'),
      body, Buffer.from('\nendstream', 'latin1'),
    ]);
  }

  addFont(ttf) {
    const f = new Font(this, ttf, `F${this.fontList.length + 1}`);
    this.fontList.push(f);
    return f;
  }

  /** 같은 그림 데이터는 한 번만 넣는다 */
  addImage(data, type) {
    if (this.imageCache.has(data)) return this.imageCache.get(data);
    const id = this.alloc();
    const img = { id, name: `Im${this.imageCache.size + 1}` };
    if (type === 'jpeg') {
      const info = jpegInfo(data);
      const cs = info.components === 1 ? '/DeviceGray' : info.components === 4 ? '/DeviceCMYK' : '/DeviceRGB';
      const decode = info.components === 4 && info.adobe ? ' /Decode [1 0 1 0 1 0 1 0]' : '';
      // JPEG은 이미 압축되어 있으므로 DCTDecode 그대로 넣는다
      this.objects[id] = Buffer.concat([
        Buffer.from(`<< /Length ${data.length} /Type /XObject /Subtype /Image /Width ${info.width} /Height ${info.height} /ColorSpace ${cs} /BitsPerComponent 8 /Filter /DCTDecode${decode} >>\nstream\n`, 'latin1'),
        data, Buffer.from('\nendstream', 'latin1'),
      ]);
      Object.assign(img, { width: info.width, height: info.height });
    } else {
      let png = encodedPng.get(data);
      if (!png) {
        const raw = decodePng(data);
        png = { width: raw.width, height: raw.height, rgb: zlib.deflateSync(raw.rgb), alpha: raw.alpha && zlib.deflateSync(raw.alpha) };
        encodedPng.set(data, png);
      }
      let smask = '';
      if (png.alpha) {
        const mid = this.alloc();
        this.setStream(mid, `/Type /XObject /Subtype /Image /Width ${png.width} /Height ${png.height} /ColorSpace /DeviceGray /BitsPerComponent 8`, null, png.alpha);
        smask = ` /SMask ${mid} 0 R`;
      }
      this.setStream(id, `/Type /XObject /Subtype /Image /Width ${png.width} /Height ${png.height} /ColorSpace /DeviceRGB /BitsPerComponent 8${smask}`, null, png.rgb);
      Object.assign(img, { width: png.width, height: png.height });
    }
    this.imageCache.set(data, img);
    return img;
  }

  addPage(width, height) {
    const p = new Page(this, width, height);
    this.pages.push(p);
    return p;
  }

  toBuffer() {
    const catalog = this.alloc(), pagesId = this.alloc(), infoId = this.alloc();
    for (const f of this.fontList) f.finish();
    const kidsIds = [];
    for (const p of this.pages) {
      const pid = this.alloc(), cid = this.alloc();
      kidsIds.push(pid);
      this.setStream(cid, '', Buffer.from(p.ops.join('\n'), 'latin1'));
      const fonts = [...p.fonts].map(f => `/${f.name} ${f.ids.type0} 0 R`).join(' ');
      const imgs = [...p.images].map(i => `/${i.name} ${i.id} 0 R`).join(' ');
      this.set(pid, `<< /Type /Page /Parent ${pagesId} 0 R /MediaBox [0 0 ${num(p.width)} ${num(p.height)}] ` +
        `/Resources << /Font << ${fonts} >> /XObject << ${imgs} >> /ProcSet [/PDF /Text /ImageB /ImageC] >> /Contents ${cid} 0 R >>`);
    }
    this.set(pagesId, `<< /Type /Pages /Kids [${kidsIds.map(i => `${i} 0 R`).join(' ')}] /Count ${kidsIds.length} >>`);
    this.set(catalog, `<< /Type /Catalog /Pages ${pagesId} 0 R >>`);
    const info = [`/Producer ${pdfString('공문 자동 작성')}`];
    if (this.info.title) info.push(`/Title ${pdfString(this.info.title)}`);
    this.set(infoId, `<< ${info.join(' ')} >>`);

    const chunks = [Buffer.from('%PDF-1.7\n%\xe2\xe3\xcf\xd3\n', 'latin1')];
    let offset = chunks[0].length;
    const offsets = [];
    for (let i = 1; i < this.objects.length; i++) {
      const body = this.objects[i] || Buffer.from('null', 'latin1');
      const obj = Buffer.concat([Buffer.from(`${i} 0 obj\n`, 'latin1'), body, Buffer.from('\nendobj\n', 'latin1')]);
      offsets[i] = offset;
      chunks.push(obj);
      offset += obj.length;
    }
    let xref = `xref\n0 ${this.objects.length}\n0000000000 65535 f \n`;
    for (let i = 1; i < this.objects.length; i++) xref += `${String(offsets[i]).padStart(10, '0')} 00000 n \n`;
    xref += `trailer\n<< /Size ${this.objects.length} /Root ${catalog} 0 R /Info ${infoId} 0 R >>\nstartxref\n${offset}\n%%EOF\n`;
    chunks.push(Buffer.from(xref, 'latin1'));
    return Buffer.concat(chunks);
  }
}

module.exports = { PdfDocument };
