'use strict';
// 브라우저용: PDF 글꼴 이름표(6글자)와 양식 구분에만 쓰는 간단한 해시(FNV-1a)
function createHash() {
  const parts = [];
  return {
    update(x) { parts.push(typeof x === 'string' ? Buffer.from(x, 'utf8') : Buffer.from(x)); return this; },
    digest(enc) {
      const data = Buffer.concat(parts), out = Buffer.alloc(32);
      for (let k = 0; k < 8; k++) {
        let h = (0x811c9dc5 ^ (k * 0x9e3779b9)) >>> 0;
        for (let i = 0; i < data.length; i++) { h ^= data[i]; h = Math.imul(h, 0x01000193) >>> 0; }
        out.writeUInt32BE(h, k * 4);
      }
      return enc ? out.toString(enc) : out;
    },
  };
}
module.exports = { createHash };
