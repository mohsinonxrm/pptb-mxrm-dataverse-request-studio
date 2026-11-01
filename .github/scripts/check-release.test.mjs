// @vitest-environment node
import { afterEach, describe, expect, it } from 'vitest';
import { spawnSync } from 'node:child_process';
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

// check-release.mjs as release.yml and CI run it (#372), in a folder with just the files it reads.
const script = fileURLToPath(new URL('./check-release.mjs', import.meta.url));
const INSIDER = '<svg>insider</svg>';
const RELEASE = '<svg>release</svg>';

let dir;
afterEach(() => dir && rmSync(dir, { recursive: true, force: true }));

function setUp({ version = '1.5.0-beta.2', bundled = INSIDER } = {}) {
  dir = mkdtempSync(join(tmpdir(), 'drs-release-check-'));
  const write = (path, text) => {
    mkdirSync(join(dir, path, '..'), { recursive: true });
    writeFileSync(join(dir, path), text);
  };
  write('package.json', JSON.stringify({ version, features: { minAPI: '1.2.0' } }));
  write('package-lock.json', JSON.stringify({ version, packages: { '': { version } } }));
  write('CHANGELOG.md', `## [Unreleased]\n\n## [${version}] - 2026-11-01\n\n- **A fix.**\n`);
  write('icons/icon-insider.svg', INSIDER);
  write('icons/icon.svg', RELEASE);
  if (bundled) write('dist/icons/icon.svg', bundled);
}

const run = (...args) => {
  const r = spawnSync(process.execPath, [script, ...args], { cwd: dir, encoding: 'utf8' });
  return { status: r.status, out: r.stdout + r.stderr };
};

describe('check-release.mjs', () => {
  it('passes, checks the icon and writes the notes', () => {
    setUp();
    const r = run('--dist', '--notes', 'notes.md', '--repository', 'o/r');
    expect(r).toMatchObject({ status: 0 });
    expect(r.out).toMatch(/matches 1\.5\.0-beta\.2, with the right icon/);
    const notes = readFileSync(join(dir, 'notes.md'), 'utf8');
    expect(notes).toMatch(/^> \*\*Requires Power Platform ToolBox 1\.2\.0/);
    expect(notes).toContain('- **A fix.**');
    expect(notes).toContain('https://github.com/o/r/blob/v1.5.0-beta.2/CHANGELOG.md');
  });

  it('fails when a stable version bundles the Insider icon', () => {
    setUp({ version: '1.5.0', bundled: INSIDER });
    const r = run('--dist');
    expect(r.status).toBe(1);
    expect(r.out).toMatch(/::error::dist\/icons\/icon\.svg isn't icons\/icon\.svg/);
  });

  it('fails --dist before the build', () => {
    setUp({ bundled: null });
    expect(run().status).toBe(0);
    const r = run('--dist');
    expect(r.status).toBe(1);
    expect(r.out).toMatch(/dist\/icons\/icon\.svg is missing/);
  });

  it('fails --notes without --repository, writing nothing', () => {
    setUp();
    const r = run('--dist', '--notes', 'notes.md');
    expect(r.status).toBe(1);
    expect(r.out).toMatch(/--notes needs --repository/);
    expect(() => readFileSync(join(dir, 'notes.md'))).toThrow();
  });
});
