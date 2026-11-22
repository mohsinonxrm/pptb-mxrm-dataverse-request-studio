#!/usr/bin/env node
// Writes assets/badges/*.svg from package.json (`npm run badges`, #373), and points the
// README's badge lines and package.json's `configurations.readmeUrl` at the version's tag (#435).
// Run it after changing the version or a badge's dependency; check-release.mjs fails when any of
// them is out of date.
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { badgeMarkdown, badges, readmeUrlFor, repositoryOf } from './badges.mjs';

const pkgText = readFileSync('package.json', 'utf8');
const pkg = JSON.parse(pkgText);
mkdirSync('assets/badges', { recursive: true });
for (const { file, svg } of badges(pkg)) {
  writeFileSync(`assets/badges/${file}`, svg);
  console.log(`assets/badges/${file}`);
}

// README: each badge line, whatever ref it points at now, becomes this version's line
const escape = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const repo = escape(repositoryOf(pkg));
const lines = badgeMarkdown(pkg);
let readme = readFileSync('README.md', 'utf8');
badges(pkg).forEach(({ file }, i) => {
  const line = new RegExp(
    `^!\\[[^\\]]*\\]\\(https://raw\\.githubusercontent\\.com/${repo}/[^)]*/assets/badges/${escape(file)}\\)$`,
    'm',
  );
  readme = readme.replace(line, lines[i]);
});
writeFileSync('README.md', readme);
console.log('README.md badge lines');

// package.json: only the readmeUrl value changes, so the file keeps its formatting
const readmeUrl = readmeUrlFor(pkg);
const nextPkg = pkgText.replace(/("readmeUrl":\s*)"[^"]*"/, `$1${JSON.stringify(readmeUrl)}`);
if (nextPkg !== pkgText) writeFileSync('package.json', nextPkg);
console.log(`package.json readmeUrl → ${readmeUrl}`);
