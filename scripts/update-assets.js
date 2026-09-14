const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');

const root = path.resolve(__dirname, '..');
const { version } = JSON.parse(fs.readFileSync(path.join(__dirname, 'release-version.json'), 'utf8'));
if (!/^[A-Za-z0-9][A-Za-z0-9._-]*$/.test(version)) throw new Error('Invalid release asset version.');

const entryAssets = ['vendor/xlsx.full.min.js', 'attachment-export.js', 'dataset-client.js', 'dataset-ui.js', 'app.js'];
const delayedAssets = ['vendor/exceljs.min.js', 'vendor/pdf-lib.min.js', 'vendor/fontkit.umd.min.js', 'vendor/open-huninn-data.js'];
const datasetWorkerAssets = [
  'vendor/zip-no-worker-inflate-2.7.57.min.js',
  'vendor/sax-1.4.1.js',
  'cdr-model.js',
  'streaming-xlsx.js',
  'dataset-store.js',
  'dataset-report.js',
  'vendor/xlsx.full.min.js',
  'attachment-export.js',
  'app.js',
  ...delayedAssets,
];
const runtimeTextAssets = [
  'index.html',
  'styles.css',
  'app.js',
  'attachment-export.js',
  'import-parser.js',
  'cdr-model.js',
  'streaming-xlsx.js',
  'dataset-client.js',
  'dataset-worker.js',
  'dataset-store.js',
  'dataset-ui.js',
  'dataset-report.js',
  ...new Set([...entryAssets, ...delayedAssets, ...datasetWorkerAssets]),
];

const assetPath = name => path.join(root, name);
const normalizeLf = text => text.replace(/\r\n?/g, '\n');
const escapeRegExp = text => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const readText = name => fs.readFileSync(assetPath(name), 'utf8');
const writeText = (name, text) => fs.writeFileSync(assetPath(name), normalizeLf(text));
const hash = name => 'sha384-' + crypto.createHash('sha384').update(fs.readFileSync(assetPath(name))).digest('base64');

function replaceRequired(text, pattern, replacement, label) {
  if (!pattern.test(text)) throw new Error(`Missing ${label}.`);
  pattern.lastIndex = 0;
  return text.replace(pattern, replacement);
}

function versionUrls(text, names, label) {
  let updated = text;
  for (const name of names) {
    const pattern = new RegExp(`(\\./${escapeRegExp(name)})(?:\\?v=[^"']*)?`, 'g');
    if (!pattern.test(updated)) throw new Error(`Missing ${label} URL for ${name}.`);
    pattern.lastIndex = 0;
    updated = updated.replace(pattern, `$1?v=${version}`);
  }
  return updated;
}

for (const name of new Set(runtimeTextAssets)) {
  if (fs.existsSync(assetPath(name))) writeText(name, readText(name));
}

const delayedHashes = new Map(delayedAssets.map(name => [name, hash(name)]));
let app = readText('app.js');
app = replaceRequired(
  app,
  /(const RELEASE_ASSET_VERSION\s*=\s*)["'][^"']+["']/,
  `$1"${version}"`,
  'app release version',
);
for (const name of delayedAssets) {
  const escapedName = escapeRegExp(name);
  const pattern = new RegExp(`(src:\\s*["']\\./${escapedName})(?:\\?v=[^"']*)?(["']\\s*,\\s*integrity:\\s*)["'][^"']*["']`);
  app = replaceRequired(app, pattern, `$1?v=${version}$2"${delayedHashes.get(name)}"`, `delayed asset declaration for ${name}`);
}
writeText('app.js', app);

let importParser = readText('import-parser.js');
importParser = versionUrls(importParser, ['vendor/xlsx.full.min.js', 'app.js'], 'import worker');
writeText('import-parser.js', importParser);

let datasetClient = readText('dataset-client.js');
datasetClient = replaceRequired(
  datasetClient,
  /(const VERSION\s*=\s*)["'][^"']+["']/,
  `$1'${version}'`,
  'dataset worker version',
);
writeText('dataset-client.js', datasetClient);

let datasetWorker = readText('dataset-worker.js');
datasetWorker = versionUrls(datasetWorker, datasetWorkerAssets, 'dataset worker');
writeText('dataset-worker.js', datasetWorker);

const entryHashes = new Map(entryAssets.map(name => [name, hash(name)]));
let html = readText('index.html');
html = replaceRequired(
  html,
  /(href=["']\.\/styles\.css)(?:\?v=[^"']*)?(["'])/,
  `$1?v=${version}$2`,
  'stylesheet loader',
);
for (const name of entryAssets) {
  const escapedName = escapeRegExp(name);
  const tagPattern = new RegExp(`<script\\b[^>]*\\bsrc=["']\\./${escapedName}(?:\\?v=[^"']*)?["'][^>]*>`);
  html = replaceRequired(html, tagPattern, (tag) => {
    const withVersion = tag.replace(
      new RegExp(`(src=["']\\./${escapedName})(?:\\?v=[^"']*)?(["'])`),
      `$1?v=${version}$2`,
    );
    return replaceRequired(withVersion, /integrity=["'][^"']*["']/, `integrity="${entryHashes.get(name)}"`, `entry integrity for ${name}`);
  }, `entry script for ${name}`);
}
const cspHashes = [...entryAssets, ...delayedAssets]
  .map(name => `'${entryHashes.get(name) || delayedHashes.get(name)}'`)
  .join(' ');
html = replaceRequired(html, /script-src 'self'[^;]*;/, `script-src 'self' ${cspHashes};`, 'script CSP directive');
writeText('index.html', html);

console.log(`Updated same-origin assets to ${version} with matching integrity hashes.`);
