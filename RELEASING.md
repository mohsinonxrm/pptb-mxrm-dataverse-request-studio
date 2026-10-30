# Releasing Dataverse Request Studio

The npm package `@mohsinonxrm/pptb-dataverse-request-studio` is published by CI
(`.github/workflows/release.yml`) when a `v*` tag is pushed. Nobody runs `npm publish`
locally. FetchXML Studio and Command Bar Studio use the same process.

## Branches

Three lanes can be live at once: the stable release, the next release in beta, and the
release after that in development. Each has its own branch.

| Branch | Holds | Tags | npm |
|---|---|---|---|
| `main` | The current **stable** release, plus hotfixes | `vX.Y.Z` | `latest` |
| `release/X.Y` | The next minor release while it **stabilizes**; exists from the feature freeze until `X.Y.0` ships | `vX.Y.0-beta.N`, `vX.Y.0-rc.N` | `beta` |
| `dev` | Development of the release after that; **never frozen** | none | — |
| `support/X.Y` | Critical fixes for the previous minor release; created only when needed (see *Supported versions*) | `vX.Y.Z` | `support-X.Y` |

For example, with 1.4 stable, 1.5 in beta and 1.6 in development: `main` is 1.4.x,
`release/1.5` tags `v1.5.0-beta.N`, and 1.6 PRs merge into `dev`.

- Work happens on branches cut from `dev` (`fix/<n>-…`, `feat/<n>-…`, `chore/…`), merged
  into `dev` by PR.
- **Every merge uses a merge commit**, never squash. That includes PRs and merges between
  the long-lived branches (`main`, `release/*`, `support/*`, `dev`), so history is shared
  and fixes are never lost.
- CI (`ci.yml`) runs on every PR and on pushes to `dev`, `main`, `release/**` and
  `support/**`. The repository ruleset requires PRs and the `check` status on every
  branch.

## Every PR

- Add your user-facing changes under `## [Unreleased]` in `CHANGELOG.md`, and update the
  README's feature docs if behaviour changes.
- Don't change the version: it's set when a release is cut. That keeps PRs independent,
  so they can merge in any order.
- Put `Fixes #N` in the PR description. `dev` is the default branch, so PRs into `dev`
  close the issue on merge. For PRs into `main`, `release/*` or `support/*`, the issue
  closes when the change reaches `dev` through the forward merge; close it by hand if you
  can't wait.
- Give the PR the milestone of the release that ships it, and an assignee.

## Dependency updates

`.github/dependabot.yml` is the same policy as FetchXML Studio's and Command Bar Studio's:

- Every Monday, Dependabot opens PRs into `dev` for npm packages and GitHub Actions.
- A new version is proposed only once it has been **published for 14 days** (cooldown).
- Minor and patch updates are grouped: Fluent UI, Monaco, `@pptb/*`, dnd-kit, build
  tooling, and all Actions. A **major arrives as its own PR**, so it's upgraded
  deliberately.
- Majors of React 18 and of `eslint` / `@eslint/js` are ignored.
  - React 18 is intentional, as in the other studios.
  - eslint 10 waits for `eslint-plugin-jsx-a11y`, which peer-requires eslint 9 or lower.
- **Security updates** come from the repository's Dependabot security updates setting and
  skip the cooldown.

Dependabot PRs follow *Every PR*, except that they don't need a CHANGELOG entry. When an
update changes the shipped bundle for a reason users care about (a security fix in a
runtime dependency), add one under *Security* in the release-prep PR. Delete the branch
after merging or closing; the repository doesn't delete head branches automatically.

## Where a fix goes

**Fix on the oldest supported branch that has the bug, then merge forward:**
`support/X.Y` → `main` → `release/X.Y` → `dev`. Never copy a fix backwards from `dev`.

| The bug exists in | Fix PR into | Release | Then merge forward |
|---|---|---|---|
| The current stable release (and later) | `hotfix/X.Y.Z` → `main` | tag `vX.Y.Z` on `main` → `latest` | `main` → `release/*` (if open) → `dev` |
| Only the release being stabilized | `release/X.Y` | tag `vX.Y.0-beta.N` → `beta` | `release/X.Y` → `dev` |
| Only `dev` | `dev` | with the next release | — |
| The previous minor only (support window) | `support/X.Y` | tag `vX.Y.Z` → `support-X.Y` | to `main` if it applies there too |

Forward merges are PRs too (`merge/release-1.5-into-dev`, merge commit). One after a batch
of beta fixes and one before the stable release is enough; they don't have to follow
every beta. When a forward merge conflicts badly because the code has moved on, re-apply
the fix on the newer branch in a small companion PR instead.

**Resolving forward-merge conflicts** in `CHANGELOG.md` and `package.json`: keep the
target branch's version and its `[Unreleased]` section on top, and add the incoming dated
sections in version order.

## Release lifecycle

### 1. Feature freeze: cut the release branch

Once the release's scope is in `dev`:

```bash
git switch dev && git pull && git switch -c release/1.5 && git push -u origin release/1.5
```

`dev` carries on with the release after that straight away. Open a **"Release X.Y.0"
tracking issue** in the milestone for the stabilizing checklist.

### 2. Betas (from `release/X.Y`)

1. On a branch from `release/1.5` (`chore/release-1.5.0-beta.1`):
   `npm version 1.5.0-beta.1 --no-git-tag-version`, rename `## [Unreleased]` to
   `## [1.5.0-beta.1] - YYYY-MM-DD` with its own `[#n]: …` link definitions (add a new
   empty `## [Unreleased]` above it), update the README if needed, and PR it into
   `release/1.5`.
2. Tag the merged commit on `release/1.5` and push the tag:

   ```bash
   git switch release/1.5 && git pull && git tag v1.5.0-beta.1 && git push origin v1.5.0-beta.1
   ```

3. Fixes found in the beta go into `release/1.5` by PR; ship them as `beta.2`, `beta.3`, ….
   Merge `release/1.5` into `dev` from time to time (see *Where a fix goes*).

`-beta.N` and `-rc.N` publish to the npm `beta` tag as a GitHub pre-release; `latest`
doesn't change. Pre-releases bundle the gold Insider icon.

### 3. Stabilizing (beta → stable)

Before promoting a beta to `latest`:

1. **Milestone clear:** every issue in the milestone is closed or moved to a later one.
2. **Regression pass and smoke test** of the latest beta, installed from npm, in Power
   Platform ToolBox against a live environment. Run the manual checks the tracking issue
   lists (header consent, dark theme, exports and imports).
3. **Soak:** a few days with no new P0/P1 bug against the latest beta. Each fix ships as a
   new `beta.N`.
4. **Health:** CI green on `release/X.Y`; `npm audit --omit=dev --audit-level=high`
   reports 0.

### 4. Stable release (to `main`)

1. Release-prep PR into `release/1.5`: `npm version 1.5.0 --no-git-tag-version`, a dated
   `## [1.5.0]` CHANGELOG section that summarises the betas above their history, and
   README updates.
2. PR `release/1.5` → `main`, merged with a merge commit.
3. Tag the merge commit on `main` and push it:

   ```bash
   git switch main && git pull && git tag v1.5.0 && git push origin v1.5.0
   ```

4. Check `npm view @mohsinonxrm/pptb-dataverse-request-studio dist-tags` and the GitHub
   Release, and smoke-test the published package.
5. Merge `main` into `dev` (PR, merge commit), then delete `release/1.5`.
6. Deprecate what the release supersedes (see *Deprecating on npm*).

## Hotfixes (to the current stable release)

1. Branch `hotfix/1.5.1` from `main`; fix, set the version (`1.5.1`), date its CHANGELOG
   section; PR into `main` (merge commit).
2. Tag `v1.5.1` on `main` → npm `latest`.
3. Merge `main` forward into `release/*` (if one is open) and `dev`, keeping each target's
   version and `[Unreleased]` entries.

## Supported versions

- **The current stable release** (`main`) gets fixes as hotfixes.
- **The previous minor release** gets **critical and security fixes for 30 days** after
  the next minor ships, for users who can't update Power Platform ToolBox yet. When such a
  fix is needed, create `support/1.4` from the last `v1.4.*` tag
  (`git switch -c support/1.4 v1.4.0 && git push -u origin support/1.4`), PR the fix into
  it, and tag `v1.4.1` there. It publishes to the npm `support-1.4` tag, never `latest`,
  and isn't marked as the latest GitHub Release.
- Older versions aren't supported.

## Deprecating on npm

Deprecate, never unpublish. Run these yourself (they need your npm login):

- When a newer beta ships, deprecate the betas it replaces:

  ```bash
  npm deprecate @mohsinonxrm/pptb-dataverse-request-studio@1.5.0-beta.1 "Superseded by 1.5.0-beta.2. Install @beta."
  ```

- When a stable release ships, deprecate the betas of that line and the previous stable
  once it's out of support:

  ```bash
  npm deprecate @mohsinonxrm/pptb-dataverse-request-studio@1.4.0 "Superseded by 1.5.0. Install @latest."
  ```

## Milestones and tracking

- Each release has a milestone (`v1.5.0`, `v1.5.1`). It's the release that **ships** the
  change, and its closed issues list what the release contains.
- What each beta contains is its `[X.Y.Z-beta.N]` CHANGELOG section and its GitHub
  pre-release notes.
- The **"Release X.Y.0" tracking issue** holds the stabilizing checklist. The milestone
  closes when the stable release ships.

## What `release.yml` does

It checks that the tag matches the `package.json` version, then chooses the channel from
the version and the tagged commit:

| Tag | Tagged commit must be on | npm tag | GitHub Release |
|---|---|---|---|
| `vX.Y.Z-beta.N`, `vX.Y.Z-rc.N` | any branch (normally `release/X.Y`) | `beta` | pre-release |
| `vX.Y.Z` | `main` | `latest` | latest |
| `vX.Y.Z` | `support/X.Y` | `support-X.Y` | not marked latest |

Any other tag fails before publishing, so a stray tag can't replace `latest`. It then:

1. runs `npm ci`, `npm test`, `npm run validate` and `npm run build`, and checks the
   bundled DOMPurify version;
2. publishes with `npm publish --provenance` through npm **trusted publishing** (OIDC, no
   npm token);
3. creates the GitHub Release. Its notes are the version's `CHANGELOG.md` section (with
   its link definitions), or generated notes when there isn't one.

Pre-release versions (anything with a `-`) bundle the gold Insider icon; stable versions
use the regular icon (`vite.config.ts`).

## One-time setup

- **npm trusted publisher:** npmjs.com → package → *Settings* → *Trusted publishing* →
  *GitHub Actions*: user `mohsinonxrm`, repository `pptb-mxrm-dataverse-request-studio`,
  workflow `release.yml`, no environment. Until it's set, the publish step fails.
- `repository.url` in `package.json` must point at this GitHub repository, or provenance
  and trusted publishing fail.
- A ruleset requiring PRs and the `check` status on `dev`, `main`, `release/**` and
  `support/**`.
