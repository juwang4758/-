'use strict';
/**
 * 양식 속 그림(PNG·JPEG)을 PDF에 넣을 수 있게 준비한다 (외부 패키지 없음)
 *  - PNG: 풀어서 RGB와 투명도(알파)를 따로 만든다. PDF에서는 알파를 SMask로 붙인다.
 *  - JPEG: 그대로 넣는다(DCTDecode). 크기와 색 공간만 읽는다.
 */
const zlib = require('zlib');

function bad(message) { return new Error(message); }

function paeth(a, b, c) {
  const p = a + b - c, pa = Math.abs(p - a), pb = Math.abs(p - b), pc = Math.abs(p - c);
  return pa <= pb && pa <= pc ? a : pb <= pc ? b : c;
}

function decodePng(buf) {
  if (buf.readUInt32BE(0) !== 0x89504e47) throw bad('PNG 파일이 아닙니다.');
  let p = 8, ihdr = null, palette = null, trns = null;
  const idat = [];
  while (p + 8 <= buf.length) {
    const len = buf.readUInt32BE(p), type = buf.toString('latin1', p + 4, p + 8);
    const data = buf.subarray(p + 8, p + 8 + len);
    if (type === 'IHDR') ihdr = { w: data.readUInt32BE(0), h: data.readUInt32BE(4), depth: data[8], color: data[9], interlace: data[12] };
    else if (type === 'PLTE') palette = data;
    else if (type === 'tRNS') trns = data;
    else if (type === 'IDAT') idat.push(data);
    else if (type === 'IEND') break;
    p += 12 + len;
  }
  if (!ihdr) throw bad('PNG 머리글이 없습니다.');
  const { w, h, depth, color } = ihdr;
  if (ihdr.interlace) throw bad('인터레이스 PNG는 지원하지 않습니다.');
  if (w * h > 25e6) throw bad('그림이 너무 큽니다.');
  const channels = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 }[color];
  if (!channels) throw bad('지원하지 않는 PNG 색 형식입니다.');
  const bitsPP = channels * depth;
  const bpp = Math.max(1, bitsPP >> 3);
  const stride = Math.ceil((w * bitsPP) / 8);
  const raw = zlib.inflateSync(Buffer.concat(idat), { maxOutputLength: (stride + 1) * h + 16 });
  const px = Buffer.alloc(stride * h);
  let prev = Buffer.alloc(stride);
  for (let y = 0; y < h; y++) {
    const f = raw[y * (stride + 1)];
    const line = raw.subarray(y * (stride + 1) + 1, (y + 1) * (stride + 1));
    const out = px.subarray(y * stride, (y + 1) * stride);
    for (let x = 0; x < stride; x++) {
      const a = x >= bpp ? out[x - bpp] : 0, b = prev[x], c = x >= bpp ? prev[x - bpp] : 0;
      const v = line[x];
      out[x] = (f === 0 ? v : f === 1 ? v + a : f === 2 ? v + b : f === 3 ? v + ((a + b) >> 1) : v + paeth(a, b, c)) & 0xff;
    }
    prev = out;
  }
  // 채널 값 하나를 원래 비트 수 그대로 읽는다(1·2·4·8·16비트)
  const raw16 = (row, idx) => {
    if (depth === 8) return px[row * stride + idx];
    if (depth === 16) return px.readUInt16BE(row * stride + idx * 2);
    const bitPos = idx * depth, byte = px[row * stride + (bitPos >> 3)];
    return (byte >> (8 - depth - (bitPos & 7))) & ((1 << depth) - 1);
  };
  const max = (1 << depth) - 1;
  const to8 = v => (depth === 8 ? v : depth === 16 ? v >> 8 : Math.round((v * 255) / max));
  const rgb = Buffer.alloc(w * h * 3);
  const alpha = Buffer.alloc(w * h, 255);
  let hasAlpha = false;
  const key = trns && (color === 0 ? [trns.readUInt16BE(0)] : color === 2 ? [trns.readUInt16BE(0), trns.readUInt16BE(2), trns.readUInt16BE(4)] : null);
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const i = y * w + x;
      let r, g, b, a = 255;
      if (color === 3) {
        const k = raw16(y, x);
        r = palette ? palette[k * 3] : 0; g = palette ? palette[k * 3 + 1] : 0; b = palette ? palette[k * 3 + 2] : 0;
        if (trns && k < trns.length) a = trns[k];
      } else if (color === 0 || color === 4) {
        const v = raw16(y, x * channels);
        r = g = b = to8(v);
        if (color === 4) a = to8(raw16(y, x * channels + 1));
        else if (key && v === key[0]) a = 0;
      } else {
        const vr = raw16(y, x * channels), vg = raw16(y, x * channels + 1), vb = raw16(y, x * channels + 2);
        r = to8(vr); g = to8(vg); b = to8(vb);
        if (color === 6) a = to8(raw16(y, x * channels + 3));
        else if (key && vr === key[0] && vg === key[1] && vb === key[2]) a = 0;
      }
      rgb[i * 3] = r; rgb[i * 3 + 1] = g; rgb[i * 3 + 2] = b;
      alpha[i] = a;
      if (a !== 255) hasAlpha = true;
    }
  }
  return { width: w, height: h, rgb, alpha: hasAlpha ? alpha : null };
}

function jpegInfo(buf) {
  if (buf[0] !== 0xff || buf[1] !== 0xd8) throw bad('JPEG 파일이 아닙니다.');
  let p = 2, adobe = false;
  while (p + 4 <= buf.length) {
    if (buf[p] !== 0xff) { p++; continue; }
    const marker = buf[p + 1];
    if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) { p += 2; continue; }
    const len = buf.readUInt16BE(p + 2);
    if (marker === 0xee && buf.toString('latin1', p + 4, p + 9) === 'Adobe') adobe = true;
    if (marker >= 0xc0 && marker <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(marker)) {
      return { height: buf.readUInt16BE(p + 5), width: buf.readUInt16BE(p + 7), components: buf[p + 9], adobe };
    }
    p += 2 + len;
  }
  throw bad('JPEG 크기를 읽을 수 없습니다.');
}

module.exports = { decodePng, jpegInfo };
