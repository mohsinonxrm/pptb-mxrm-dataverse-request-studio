import { describe, expect, it } from 'vitest';
import {
  changelogSections,
  checkReleaseFiles,
  expectedIconSource,
  releaseNotes,
} from './release-checks.mjs';

const pkg = {
  name: '@scope/tool',
  displayName: 'Tool Studio',
  version: '1.5.0-beta.2',
  features: { minAPI: '1.2.0' },
};
const lock = { version: '1.5.0-beta.2', packages: { '': { version: '1.5.0-beta.2' } } };
const changelog = `# Changelog

## [Unreleased]

## [1.5.0-beta.2] - 2026-11-01

The second beta.

### 🐛 Fixed

- **A fix** ([#69]).

[#69]: https://github.com/o/r/issues/69

## [1.5.0-beta.1] - 2026-10-30

- Older.
`;
const files = { pkg, lock, changelog };

describe('checkReleaseFiles', () => {
  it('passes when everything matches the version', () => {
    expect(checkReleaseFiles(files)).toEqual([]);
  });

  it('allows entries under [Unreleased] above the release', () => {
    const withPending = changelog.replace(
      '## [Unreleased]\n',
      '## [Unreleased]\n\n### 🐛 Fixed\n\n- **Pending.**\n',
    );
    expect(checkReleaseFiles({ ...files, changelog: withPending })).toEqual([]);
  });

  it("flags a version release.yml doesn't publish", () => {
    const problems = checkReleaseFiles({ ...files, pkg: { ...pkg, version: '1.5.0-alpha.1' } });
    expect(problems[0]).toMatch(/isn't X\.Y\.Z/);
  });

  it('flags package-lock.json left on the old version', () => {
    const problems = checkReleaseFiles({
      ...files,
      lock: { ...lock, packages: { '': { version: '1.5.0-beta.1' } } },
    });
    expect(problems).toEqual([
      expect.stringMatching(/^package-lock\.json has 1\.5\.0-beta\.2 \/ 1\.5\.0-beta\.1/),
    ]);
  });

  it("flags a CHANGELOG whose newest section isn't the version (a forgotten rename)", () => {
    const stale = changelog.replace('## [1.5.0-beta.2] - 2026-11-01', '');
    expect(checkReleaseFiles({ ...files, changelog: stale }).join(' ')).toMatch(
      /newest release section is \[1\.5\.0-beta\.1\], not \[1\.5\.0-beta\.2\]/,
    );
  });

  it('flags a missing [Unreleased], a missing or invalid date and an empty section', () => {
    const check = (cl) => checkReleaseFiles({ ...files, changelog: cl }).join(' ');
    expect(check(changelog.replace('## [Unreleased]\n', ''))).toMatch(/must start with/);
    expect(check(changelog.replace(' - 2026-11-01', ''))).toMatch(/needs a date/);
    expect(check(changelog.replace('2026-11-01', '2026-02-30'))).toMatch(/needs a date/);
    const empty = changelog.replace('### 🐛 Fixed\n\n- **A fix** ([#69]).\n', '');
    expect(check(empty)).toMatch(/has no entries/);
  });

  it('flags a duplicated section', () => {
    const twice = `${changelog}\n## [1.5.0-beta.2] - 2026-11-01\n\n- Again.\n`;
    expect(checkReleaseFiles({ ...files, changelog: twice }).join(' ')).toMatch(
      /2 \[1\.5\.0-beta\.2\] sections/,
    );
  });

  it('flags a second [Unreleased] a forward merge left below a release', () => {
    const stranded = changelog.replace(
      '## [1.5.0-beta.1]',
      '## [Unreleased]\n\n- **Stranded.**\n\n## [1.5.0-beta.1]',
    );
    expect(checkReleaseFiles({ ...files, changelog: stranded })).toEqual([
      expect.stringMatching(/2 \[Unreleased\] sections/),
    ]);
  });

  it('passes a release candidate and a stable release', () => {
    for (const version of ['1.5.0-rc.1', '1.5.0']) {
      expect(
        checkReleaseFiles({
          pkg: { ...pkg, version },
          lock: { version, packages: { '': { version } } },
          changelog: changelog.replace('[1.5.0-beta.2]', `[${version}]`),
        }),
      ).toEqual([]);
    }
  });
});

describe('expectedIconSource / changelogSections', () => {
  it('bundles the Insider icon for pre-releases only', () => {
    expect(expectedIconSource('1.5.0-beta.2')).toBe('icons/icon-insider.svg');
    expect(expectedIconSource('1.5.0-rc.1')).toBe('icons/icon-insider.svg');
    expect(expectedIconSource('1.5.0')).toBe('icons/icon.svg');
  });

  it('reads the date of headings with text after it, and keeps their bodies apart', () => {
    const yanked = changelog.replace(
      '## [1.5.0-beta.1] - 2026-10-30',
      '## [1.5.0-beta.1] - 2026-10-30 [YANKED]',
    );
    const sections = changelogSections(yanked);
    expect(sections[2]).toMatchObject({
      name: '1.5.0-beta.1',
      date: '2026-10-30',
      body: '- Older.',
    });
    expect(sections[1].body).not.toContain('Older');
  });

  it('reads headings and bodies, normalizing CRLF line endings', () => {
    const sections = changelogSections(changelog.replace(/\n/g, '\r\n'));
    expect(sections.map((s) => [s.name, s.date])).toEqual([
      ['Unreleased', undefined],
      ['1.5.0-beta.2', '2026-11-01'],
      ['1.5.0-beta.1', '2026-10-30'],
    ]);
    expect(sections[1].body).toBe(
      'The second beta.\n\n### 🐛 Fixed\n\n- **A fix** ([#69]).\n\n[#69]: https://github.com/o/r/issues/69',
    );
  });
});

describe('releaseNotes', () => {
  it('leads with the PPTB requirement, then the section, keeping its [#n] link definitions', () => {
    expect(releaseNotes({ pkg, changelog, repository: 'o/r' }))
      .toBe(`> **Requires Power Platform ToolBox 1.2.0 or later** (\`minAPI\` 1.2.0).

The second beta.

### 🐛 Fixed

- **A fix** ([#69]).

[#69]: https://github.com/o/r/issues/69

See [CHANGELOG.md](https://github.com/o/r/blob/v1.5.0-beta.2/CHANGELOG.md) for earlier versions.
`);
  });

  it('leaves out the requirement without a minAPI', () => {
    const notes = releaseNotes({ pkg: { ...pkg, features: {} }, changelog, repository: 'o/r' });
    expect(notes.startsWith('The second beta.')).toBe(true);
  });

  it("fails without the version's section", () => {
    expect(() =>
      releaseNotes({ pkg: { ...pkg, version: '9.9.9' }, changelog, repository: 'o/r' }),
    ).toThrow(/no \[9\.9\.9\] section/);
  });
});
