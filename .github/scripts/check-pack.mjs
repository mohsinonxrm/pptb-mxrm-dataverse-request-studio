#!/usr/bin/env node
// Checks `npm pack --dry-run --json` output (#273): the files PPTB needs are
// in the package, nothing dev-only is, and it stays under a size budget.
// Shared with FetchXML Studio and Command Bar Studio.
//
// Usage: npm pack --dry-run --json --silent > pack.json && node .github/scripts/check-pack.mjs pack.json
import { appendFileSync, readFileSync } from 'node:fs';

// Budgets with headroom over the bundle trim (#236), which dropped Monaco's
// TS, CSS and HTML language services and ~80 unused language chunks. Before
// it: ~14.5 MB unpacked, ~3.6 MB packed, ~2 MB loaded at startup.
const MAX_UNPACKED_BYTES = 9 * 1024 * 1024; // #236: ~6.7 MB
const MAX_PACKED_BYTES = 2.5 * 1024 * 1024; // #236: ~1.8 MB
// The scripts dist/index.html loads at startup (its module script and any
// modulepreloads).
const MAX_BOOT_BYTES = 2 * 1024 * 1024; // #236: ~1.7 MB
// Files only a full `import 'monaco-editor'` (or a worker DRS doesn't use)
// would emit: its TS / CSS / HTML language services and workers, and a few
// of the ~80 basic languages it registers.
const UNUSED_MONACO =
  /^dist\/assets\/((ts|css|html)\.worker|tsMode|cssMode|htmlMode|abap|apex|cameligo|solidity)-[\w-]+\.js$/;

const REQUIRED = [
  'package.json',
  'index.html',
  'dist/index.html',
  'dist/icons/icon.svg',
  'README.md',
  'CHANGELOG.md',
  'LICENSE',
  'SECURITY.md',
];
const FORBIDDEN = [/^src\//, /^dev\//, /^coverage\//, /^\.github\//, /^\.claude\//, /\.map$/];

const raw = readFileSync(process.argv[2] ?? 0, 'utf8');
// npm <= 11 prints an array of packages, npm 12 an object keyed by package
// name. Tolerate stray lifecycle-script output around the JSON.
const start = raw.search(/[[{]/);
const end = Math.max(raw.lastIndexOf(']'), raw.lastIndexOf('}'));
const parsed = JSON.parse(raw.slice(start, end + 1));
const pkg = Array.isArray(parsed) ? parsed[0] : Object.values(parsed)[0];
const files = pkg.files.map((f) => f.path.replace(/\\/g, '/'));

const errors = [];
for (const req of REQUIRED) {
  if (!files.includes(req)) errors.push(`missing required file: ${req}`);
}
for (const file of files) {
  if (FORBIDDEN.some((re) => re.test(file))) errors.push(`forbidden file in package: ${file}`);
}
if (!files.some((f) => /^dist\/assets\/index-[\w-]+\.js$/.test(f))) {
  errors.push('missing the app bundle: dist/assets/index-*.js');
}
// Startup cost: what dist/index.html loads, read from the built file (the
// check runs at the package root, after the build).
const sizeOf = new Map(pkg.files.map((f) => [f.path.replace(/\\/g, '/'), f.size]));
let html = '';
try {
  html = readFileSync('dist/index.html', 'utf8');
} catch {
  errors.push("can't read dist/index.html to measure what loads at startup");
}
const bootFiles = [
  ...html.matchAll(/<script\b[^>]*\bsrc="([^"]+\.js)"/g),
  ...html.matchAll(/<link\b(?=[^>]*\brel="modulepreload")[^>]*\bhref="([^"]+\.js)"/g),
].map((m) => `dist/${m[1].replace(/^\.?\//, '')}`);
let bootBytes = 0;
for (const file of bootFiles) {
  const size = sizeOf.get(file);
  if (typeof size !== 'number') errors.push(`startup script ${file} isn't in the package`);
  else bootBytes += size;
}
if (html && bootFiles.length === 0) errors.push('dist/index.html loads no module script');
if (bootBytes > MAX_BOOT_BYTES) {
  errors.push(`startup scripts are ${bootBytes} B, over the ${MAX_BOOT_BYTES} B budget`);
}
for (const file of files.filter((f) => UNUSED_MONACO.test(f))) {
  errors.push(`unused Monaco code in package (a full monaco-editor import?): ${file}`);
}
if (pkg.unpackedSize > MAX_UNPACKED_BYTES) {
  errors.push(`unpacked size ${pkg.unpackedSize} B exceeds budget ${MAX_UNPACKED_BYTES} B`);
}
if (pkg.size > MAX_PACKED_BYTES) {
  errors.push(`packed size ${pkg.size} B exceeds budget ${MAX_PACKED_BYTES} B`);
}

const mb = (n) => `${(n / 1024 / 1024).toFixed(2)} MB`;
const summary = [
  `### 📦 ${pkg.name}@${pkg.version}`,
  '',
  '| files | packed | unpacked | loaded at startup |',
  '| --- | --- | --- | --- |',
  `| ${pkg.entryCount} | ${mb(pkg.size)} | ${mb(pkg.unpackedSize)} | ${mb(bootBytes)} |`,
  '',
  errors.length
    ? `❌ ${errors.length} problem(s):\n${errors.map((e) => `- ${e}`).join('\n')}`
    : '✅ package contents OK',
  '',
].join('\n');

console.log(summary);
if (process.env.GITHUB_STEP_SUMMARY) appendFileSync(process.env.GITHUB_STEP_SUMMARY, summary);
if (errors.length) process.exit(1);
