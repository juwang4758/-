'use strict';
// 브라우저용: Node zlib 대신 pako(순수 JS)로 같은 동기 함수를 제공한다
const pako = require('pako');
const out = u8 => Buffer.from(u8.buffer, u8.byteOffset, u8.byteLength);
module.exports = {
  inflateSync: d => out(pako.inflate(d)),
  inflateRawSync: d => out(pako.inflateRaw(d)),
  deflateSync: d => out(pako.deflate(d)),
  deflateRawSync: d => out(pako.deflateRaw(d)),
};
