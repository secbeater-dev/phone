const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
const root = path.resolve(__dirname, '..');
const version = '20260911-legacy-ui-v1';
const scripts = ['vendor/xlsx.full.min.js', 'attachment-export.js', 'dataset-client.js', 'dataset-ui.js', 'app.js'];
// Git checks these assets out as LF; hash those exact bytes on Windows too.
for (const name of fs.readdirSync(root).filter(name => /\.(js|css|html)$/.test(name))) {
  const file = path.join(root, name);
  fs.writeFileSync(file, fs.readFileSync(file, 'utf8').replace(/\r\n/g, '\n'));
}
let html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const hash = name => 'sha384-' + crypto.createHash('sha384').update(fs.readFileSync(path.join(root, name))).digest('base64');
for (const name of ['app.js', 'import-parser.js']) {
  const file = path.join(root, name);
  fs.writeFileSync(file, fs.readFileSync(file, 'utf8').replace(/20260827-notice-card-layout-v1/g, version).replace(/\r\n/g, '\n'));
}
html = html.replace(/    <script src="\.\/(?:vendor\/xlsx.full.min.js|attachment-export.js|dataset-client.js|dataset-ui.js|app.js)[^\n]*\n/g, '');
html = html.replace('  </body>', scripts.map(name => `    <script src="./${name}?v=${version}" integrity="${hash(name)}" crossorigin="anonymous"></script>`).join('\n') + '\n  </body>');
const hashes = scripts.concat(['vendor/exceljs.min.js','vendor/pdf-lib.min.js','vendor/fontkit.umd.min.js','vendor/open-huninn-data.js']).map(name => "'" + hash(name) + "'").join(' ');
html = html.replace(/script-src 'self'[^;]+;/, "script-src 'self' " + hashes + ';').replace(/styles.css\?v=[^" ]+/, 'styles.css?v=' + version);
fs.writeFileSync(path.join(root, 'index.html'), html.replace(/\r\n/g, '\n'));
console.log('Updated same-origin asset version and integrity hashes.');
