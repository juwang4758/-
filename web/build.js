'use strict';
// GitHub Pages용 정적 사이트 만들기: node web/build.js → site/
const fs = require('fs');
const path = require('path');
const esbuild = require('esbuild');
const root = path.join(__dirname, '..');
const out = path.join(root, 'site');
fs.rmSync(out, { recursive: true, force: true });
fs.mkdirSync(path.join(out, 'fonts'), { recursive: true });
esbuild.buildSync({
  entryPoints: [path.join(__dirname, 'local-api.js')], bundle: true, minify: true, format: 'iife', target: ['es2020'],
  outfile: path.join(out, 'app.js'), inject: [path.join(__dirname, 'shims/buffer.js')], platform: 'browser',
  alias: { zlib: path.join(__dirname, 'shims/zlib.js'), crypto: path.join(__dirname, 'shims/crypto.js'), path: 'path-browserify' },
  logLevel: 'warning',
});
const html = fs.readFileSync(path.join(root, 'public/index.html'), 'utf8').replace('<script>', '<script src="app.js"></script>\n<script>');
fs.writeFileSync(path.join(out, 'index.html'), html);
for (const f of fs.readdirSync(path.join(root, 'fonts'))) fs.copyFileSync(path.join(root, 'fonts', f), path.join(out, 'fonts', f));
fs.writeFileSync(path.join(out, '.nojekyll'), '');
console.log('site/ 완료:', fs.readdirSync(out).join(', '));
