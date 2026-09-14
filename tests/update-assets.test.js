const test = require("node:test");
const assert = require("node:assert/strict");
const crypto = require("node:crypto");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const { spawnSync } = require("node:child_process");

const SOURCE_ROOT = path.resolve(__dirname, "..");
const TARGET_VERSION = "20991231-synthetic-release-v9";
const ENTRY_ASSETS = [
  "vendor/xlsx.full.min.js",
  "attachment-export.js",
  "dataset-client.js",
  "dataset-ui.js",
  "app.js",
];
const DELAYED_ASSETS = [
  "vendor/exceljs.min.js",
  "vendor/pdf-lib.min.js",
  "vendor/fontkit.umd.min.js",
  "vendor/open-huninn-data.js",
];
const WORKER_ASSETS = [
  "vendor/zip-no-worker-inflate-2.7.57.min.js",
  "vendor/sax-1.4.1.js",
  "cdr-model.js",
  "streaming-xlsx.js",
  "dataset-store.js",
  "dataset-report.js",
  "vendor/xlsx.full.min.js",
  "attachment-export.js",
  "app.js",
  ...DELAYED_ASSETS,
];
const RUNTIME_FILES = [
  "index.html",
  "styles.css",
  "app.js",
  "attachment-export.js",
  "import-parser.js",
  "cdr-model.js",
  "streaming-xlsx.js",
  "dataset-client.js",
  "dataset-worker.js",
  "dataset-store.js",
  "dataset-ui.js",
  "dataset-report.js",
  ...new Set([...ENTRY_ASSETS, ...DELAYED_ASSETS, ...WORKER_ASSETS]),
];

function sri(file) {
  return `sha384-${crypto.createHash("sha384").update(fs.readFileSync(file)).digest("base64")}`;
}

function writeFixtureFile(root, relativePath, contents) {
  const file = path.join(root, relativePath);
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, contents);
}

function makeFixture(t) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "phone-update-assets-"));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  fs.mkdirSync(path.join(root, "scripts"), { recursive: true });
  fs.copyFileSync(path.join(SOURCE_ROOT, "scripts", "update-assets.js"), path.join(root, "scripts", "update-assets.js"));
  writeFixtureFile(root, "scripts/release-version.json", `${JSON.stringify({ version: TARGET_VERSION })}\n`);

  writeFixtureFile(root, "index.html", [
    "<!doctype html>",
    '<meta http-equiv="Content-Security-Policy" content="default-src \'none\'; script-src \'self\' \'sha384-stale\'; worker-src \'self\'">',
    '<link rel="stylesheet" href="./styles.css?v=20260911-old-release">',
    ...ENTRY_ASSETS.map((name) => `<script src="./${name}?v=20260911-old-release" integrity="sha384-stale" crossorigin="anonymous"></script>`),
  ].join("\r\n") + "\r\n");
  writeFixtureFile(root, "app.js", [
    'const RELEASE_ASSET_VERSION = "20260911-old-release";',
    'const unrelatedReleaseNote = "20260911-old-release";',
    'exceljs: { src: "./vendor/exceljs.min.js?v=20260911-old-release", integrity: "sha384-stale" },',
    'pdfLib: { src: "./vendor/pdf-lib.min.js?v=20260911-old-release", integrity: "sha384-stale" },',
    'fontkit: { src: "./vendor/fontkit.umd.min.js?v=20260911-old-release", integrity: "sha384-stale" },',
    'fontData: { src: "./vendor/open-huninn-data.js?v=20260911-old-release", integrity: "sha384-stale" },',
    'new Worker("./import-parser.js?v=" + RELEASE_ASSET_VERSION);',
  ].join("\r\n") + "\r\n");
  writeFixtureFile(root, "import-parser.js", 'importScripts("./vendor/xlsx.full.min.js?v=20260827-older", "./app.js?v=20260827-older");\r\n');
  writeFixtureFile(root, "dataset-client.js", "const VERSION = '20260910-current';\r\nnew Worker('./dataset-worker.js?v=' + VERSION);\r\n");
  writeFixtureFile(root, "dataset-worker.js", [
    "importScripts('./vendor/zip-no-worker-inflate-2.7.57.min.js', './vendor/sax-1.4.1.js');",
    "importScripts('./cdr-model.js?v=20260910-current', './streaming-xlsx.js?v=20260910-current', './dataset-store.js?v=20260910-current', './dataset-report.js?v=20260910-current');",
    "importScripts('./vendor/xlsx.full.min.js', './attachment-export.js?v=20260910-current', './app.js?v=20260910-current');",
    "importScripts('./vendor/exceljs.min.js');",
    "importScripts('./vendor/pdf-lib.min.js', './vendor/fontkit.umd.min.js', './vendor/open-huninn-data.js');",
  ].join("\r\n") + "\r\n");

  for (const relativePath of RUNTIME_FILES) {
    if (!fs.existsSync(path.join(root, relativePath))) writeFixtureFile(root, relativePath, `${relativePath}\r\nsynthetic bytes\r\n`);
  }
  return root;
}

function runUpdater(root) {
  const result = spawnSync(process.execPath, [path.join(root, "scripts", "update-assets.js")], {
    cwd: root,
    encoding: "utf8",
  });
  assert.equal(result.status, 0, result.stderr || result.stdout);
}

test("release updater upgrades the complete loader graph and hashes normalized bytes", (t) => {
  const root = makeFixture(t);

  runUpdater(root);

  const html = fs.readFileSync(path.join(root, "index.html"), "utf8");
  const app = fs.readFileSync(path.join(root, "app.js"), "utf8");
  const importParser = fs.readFileSync(path.join(root, "import-parser.js"), "utf8");
  const datasetClient = fs.readFileSync(path.join(root, "dataset-client.js"), "utf8");
  const datasetWorker = fs.readFileSync(path.join(root, "dataset-worker.js"), "utf8");

  assert.match(html, new RegExp(`styles\\.css\\?v=${TARGET_VERSION}`));
  assert.deepEqual(
    Array.from(html.matchAll(/<script src="\.\/([^?]+)\?v=[^"]+"/g), (match) => match[1]),
    ENTRY_ASSETS,
  );
  for (const name of ENTRY_ASSETS) {
    const hash = sri(path.join(root, name));
    assert.ok(html.includes(`src="./${name}?v=${TARGET_VERSION}" integrity="${hash}"`), name);
    assert.ok(html.includes(`'${hash}'`), `${name} CSP hash`);
  }
  for (const name of DELAYED_ASSETS) {
    const hash = sri(path.join(root, name));
    assert.ok(app.includes(`src: "./${name}?v=${TARGET_VERSION}", integrity: "${hash}"`), name);
    assert.ok(html.includes(`'${hash}'`), `${name} CSP hash`);
  }
  assert.match(app, new RegExp(`RELEASE_ASSET_VERSION = "${TARGET_VERSION}"`));
  assert.match(app, /unrelatedReleaseNote = "20260911-old-release"/);
  assert.deepEqual(Array.from(importParser.matchAll(/\?v=([^"']+)/g), (match) => match[1]), [TARGET_VERSION, TARGET_VERSION]);
  assert.match(datasetClient, new RegExp(`const VERSION = '${TARGET_VERSION}'`));
  for (const name of WORKER_ASSETS) {
    assert.ok(datasetWorker.includes(`./${name}?v=${TARGET_VERSION}`), name);
  }
  for (const relativePath of RUNTIME_FILES) {
    assert.doesNotMatch(fs.readFileSync(path.join(root, relativePath), "utf8"), /\r/, `${relativePath} must be LF`);
  }
});

test("release updater is byte-identical on a second run", (t) => {
  const root = makeFixture(t);
  runUpdater(root);
  const first = new Map(RUNTIME_FILES.map((name) => [name, fs.readFileSync(path.join(root, name))]));

  runUpdater(root);

  for (const [name, bytes] of first) assert.deepEqual(fs.readFileSync(path.join(root, name)), bytes, name);
});
