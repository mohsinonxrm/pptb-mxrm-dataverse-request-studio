#!/usr/bin/env node
// Checks `npm pack --dry-run --json` output (#273): the files PPTB needs are
// in the package, nothing dev-only is, and it stays under a size budget.
// Shared with FetchXML Studio and Command Bar Studio.
//
// Usage: npm pack --dry-run --json --silent > pack.json && node .github/scripts/check-pack.mjs pack.json
import { appendFileSync, readFileSync } from 'node:fs';

const MAX_UNPACKED_BYTES = 22 * 1024 * 1024; // 1.4.0: ~14.5 MB (Monaco is most of it)
const MAX_PACKED_BYTES = 6 * 1024 * 1024; // 1.4.0: ~3.6 MB

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
  '| files | packed | unpacked |',
  '| --- | --- | --- |',
  `| ${pkg.entryCount} | ${mb(pkg.size)} | ${mb(pkg.unpackedSize)} |`,
  '',
  errors.length
    ? `❌ ${errors.length} problem(s):\n${errors.map((e) => `- ${e}`).join('\n')}`
    : '✅ package contents OK',
  '',
].join('\n');

console.log(summary);
if (process.env.GITHUB_STEP_SUMMARY) appendFileSync(process.env.GITHUB_STEP_SUMMARY, summary);
if (errors.length) process.exit(1);
