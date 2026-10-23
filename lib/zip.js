'use strict';
/**
 * ZIP 읽기·쓰기 (외부 패키지 없음)
 *  - 읽기: 엑셀(.xlsx) 파일을 여는 데 쓴다. 저장(0)·deflate(8) 방식만 지원한다.
 *  - 쓰기: 교회별 PDF 여러 개를 한 번에 내려받을 때 쓴다. 파일 이름은 UTF-8로 표시한다.
 */
const zlib = require('zlib');

const MAX_ENTRY = 40 * 1024 * 1024; // 압축 폭탄 방지: 항목 하나를 풀었을 때 최대 크기

let crcTable = null;
function crc32(buf) {
  if (zlib.crc32) return zlib.crc32(buf);
  if (!crcTable) {
    crcTable = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      crcTable[n] = c >>> 0;
    }
  }
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = crcTable[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function bad(message) { return Object.assign(new Error(message), { status: 400 }); }

/** ZIP을 읽어 { 경로: () => Buffer } 를 돌려준다. 내용은 필요할 때만 푼다 */
function readZip(buf) {
  let eocd = -1;
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 65557); i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw bad('엑셀(.xlsx) 파일이 아니거나 손상되었습니다.');
  const count = buf.readUInt16LE(eocd + 10);
  let p = buf.readUInt32LE(eocd + 16);
  const entries = {};
  for (let i = 0; i < count; i++) {
    if (p + 46 > buf.length || buf.readUInt32LE(p) !== 0x02014b50) throw bad('엑셀 파일 구조를 읽을 수 없습니다.');
    const method = buf.readUInt16LE(p + 10);
    const compSize = buf.readUInt32LE(p + 20);
    const size = buf.readUInt32LE(p + 24);
    const nameLen = buf.readUInt16LE(p + 28), extraLen = buf.readUInt16LE(p + 30), commentLen = buf.readUInt16LE(p + 32);
    const local = buf.readUInt32LE(p + 42);
    const name = buf.toString('utf8', p + 46, p + 46 + nameLen);
    p += 46 + nameLen + extraLen + commentLen;
    entries[name] = () => {
      if (buf.readUInt32LE(local) !== 0x04034b50) throw bad('엑셀 파일 구조를 읽을 수 없습니다.');
      const start = local + 30 + buf.readUInt16LE(local + 26) + buf.readUInt16LE(local + 28);
      const raw = buf.subarray(start, start + compSize);
      if (method === 0) return raw;
      if (method !== 8) throw bad('지원하지 않는 압축 방식의 엑셀 파일입니다.');
      if (size > MAX_ENTRY) throw bad('엑셀 파일 안의 내용이 너무 큽니다.');
      return zlib.inflateRawSync(raw, { maxOutputLength: MAX_ENTRY });
    };
  }
  return entries;
}

function dosTime(d) {
  return {
    time: (d.getHours() << 11) | (d.getMinutes() << 5) | Math.floor(d.getSeconds() / 2),
    date: ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate(),
  };
}

/** files: [{ name, data }] → ZIP Buffer */
function writeZip(files, now = new Date()) {
  const { time, date } = dosTime(now);
  const locals = [], centrals = [];
  let offset = 0;
  for (const f of files) {
    const name = Buffer.from(f.name, 'utf8');
    const comp = zlib.deflateRawSync(f.data);
    const crc = crc32(f.data);
    const head = Buffer.alloc(30);
    head.writeUInt32LE(0x04034b50, 0); head.writeUInt16LE(20, 4); head.writeUInt16LE(0x0800, 6); head.writeUInt16LE(8, 8);
    head.writeUInt16LE(time, 10); head.writeUInt16LE(date, 12); head.writeUInt32LE(crc, 14);
    head.writeUInt32LE(comp.length, 18); head.writeUInt32LE(f.data.length, 22); head.writeUInt16LE(name.length, 26);
    const cen = Buffer.alloc(46);
    cen.writeUInt32LE(0x02014b50, 0); cen.writeUInt16LE(20, 4); cen.writeUInt16LE(20, 6); cen.writeUInt16LE(0x0800, 8);
    cen.writeUInt16LE(8, 10); cen.writeUInt16LE(time, 12); cen.writeUInt16LE(date, 14); cen.writeUInt32LE(crc, 16);
    cen.writeUInt32LE(comp.length, 20); cen.writeUInt32LE(f.data.length, 24); cen.writeUInt16LE(name.length, 28);
    cen.writeUInt32LE(offset, 42);
    locals.push(head, name, comp);
    centrals.push(cen, name);
    offset += head.length + name.length + comp.length;
  }
  const central = Buffer.concat(centrals);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(files.length, 8); end.writeUInt16LE(files.length, 10);
  end.writeUInt32LE(central.length, 12); end.writeUInt32LE(offset, 16);
  return Buffer.concat([...locals, central, end]);
}

module.exports = { readZip, writeZip, crc32 };
