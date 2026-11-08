import { describe, expect, it } from 'vitest';
import { badgeMarkdown, badges, badgeSvg, badgeUrl } from './badges.mjs';
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
  license: 'AGPL-3.0-only',
  repository: { type: 'git', url: 'git+https://github.com/owner/tool.git' },
  features: { minAPI: '1.2.0' },
  dependencies: { react: '^18.3.1', '@fluentui/react-components': '^9.73.8' },
  devDependencies: { typescript: '~5.9.3', '@pptb/types': '^1.2.7' },
};
/** The README and assets/badges/ files as `npm run badges` leaves them for a package.json */
const readmeFor = (p) => `# Tool\n\n${badgeMarkdown(p).join('\n')}\n`;
const badgeFilesFor = (p) => Object.fromEntries(badges(p).map((b) => [b.file, b.svg]));
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
const readme = readmeFor(pkg);
const badgeFiles = badgeFilesFor(pkg);
const files = { pkg, lock, changelog, readme, badgeFiles };

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
      const p = { ...pkg, version };
      expect(
        checkReleaseFiles({
          pkg: p,
          lock: { version, packages: { '': { version } } },
          changelog: changelog.replace('[1.5.0-beta.2]', `[${version}]`),
          readme: readmeFor(p),
          badgeFiles: badgeFilesFor(p),
        }),
      ).toEqual([]);
    }
  });

  it('flags badge files not regenerated for the version, or missing', () => {
    const old = badgeFilesFor({ ...pkg, version: '1.5.0-beta.1' })['version.svg'];
    expect(
      checkReleaseFiles({ ...files, badgeFiles: { ...badgeFiles, 'version.svg': old } }),
    ).toEqual(['assets/badges/ is out of date (version.svg): run `npm run badges`.']);
    expect(
      checkReleaseFiles({ ...files, badgeFiles: { ...badgeFiles, 'react.svg': undefined } }).join(
        ' ',
      ),
    ).toMatch(/out of date \(react\.svg\)/);
  });

  it('flags a README without the badges, or with img.shields.io', () => {
    expect(checkReleaseFiles({ ...files, readme: '# Tool\n' }).join(' ')).toMatch(
      /README\.md doesn't show these badges/,
    );
    const shields = `${readme}![npm](https://img.shields.io/npm/v/@scope/tool)\n`;
    expect(checkReleaseFiles({ ...files, readme: shields })).toEqual([
      expect.stringMatching(/img\.shields\.io, which PPTB's README viewer blocks/),
    ]);
    for (const image of [
      '<img src="https://img.shields.io/badge/x-y-blue">',
      '<img src = "https://img.shields.io/badge/x-y-blue">',
      '![npm][npm]\n\n[npm]: https://img.shields.io/npm/v/x',
      '![npm](<https://img.shields.io/npm/v/x>)',
    ]) {
      expect(checkReleaseFiles({ ...files, readme: `${readme}${image}\n` })).toHaveLength(1);
    }
    // A mention in the text, or a URL in code, is fine
    for (const text of [
      "The badges aren't on img.shields.io.",
      'Not `https://img.shields.io/badge/x-y-blue`.',
      '```md\n![x](https://img.shields.io/badge/x-y-blue)\n```',
    ]) {
      expect(checkReleaseFiles({ ...files, readme: `${readme}${text}\n` })).toEqual([]);
    }
  });
});

describe('badges', () => {
  it('generates every badge from package.json, linked from main on raw.githubusercontent.com', () => {
    expect(badges(pkg).map((b) => [b.file, b.alt])).toEqual([
      ['version.svg', 'Version'],
      ['pptb.svg', 'Power Platform ToolBox 1.2.0 or later'],
      ['pptb-types.svg', '@pptb/types 1.2.7'],
      ['react.svg', 'React 18'],
      ['typescript.svg', 'TypeScript 5.9'],
      ['fluent-ui.svg', 'Fluent UI v9'],
      ['license.svg', 'License AGPL-3.0-only'],
    ]);
    expect(badgeUrl(pkg, 'react.svg')).toBe(
      'https://raw.githubusercontent.com/owner/tool/refs/heads/main/assets/badges/react.svg',
    );
    // The version badge links to npm and has no version in its alt text, so a release
    // changes only the SVG
    expect(badgeMarkdown(pkg)[0]).toBe(
      '[![Version](https://raw.githubusercontent.com/owner/tool/refs/heads/main/assets/badges/version.svg)](https://www.npmjs.com/package/@scope/tool)',
    );
  });

  it('colours pre-releases yellow and releases blue, and escapes text', () => {
    const svgOf = (version) => badgeFilesFor({ ...pkg, version })['version.svg'];
    expect(svgOf('1.5.0-beta.2')).toContain('fill="#dfb317"');
    expect(svgOf('1.5.0')).toContain('fill="#007ec6"');
    expect(svgOf('1.5.0')).toContain('aria-label="version: 1.5.0"');
    expect(badgeSvg('a<b', 'c&d', '#000')).toContain('a&lt;b: c&amp;d');
  });

  it('needs a GitHub repository.url', () => {
    expect(() => badgeUrl({ ...pkg, repository: 'https://example.com/x' }, 'a.svg')).toThrow(
      /isn't a GitHub repository/,
    );
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
