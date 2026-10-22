# Releasing Dataverse Request Studio

The npm package `@mohsinonxrm/pptb-dataverse-request-studio` is published by CI
(`.github/workflows/release.yml`) when a `v*` tag is pushed. Nobody runs `npm publish`
locally. FetchXML Studio and Command Bar Studio use the same workflow.

## How a tag publishes

| Tag | npm dist-tag | GitHub Release |
|---|---|---|
| `vX.Y.Z-beta.N`, `vX.Y.Z-rc.N` | `beta` | pre-release |
| `vX.Y.Z` on `main` | `latest` | latest release |
| `vX.Y.Z` on `support/X.Y` | `support-X.Y` | release, not marked latest |

Other pre-release suffixes, and a stable tag on any other branch, fail the workflow, so a
stray tag can't replace `latest`.

The workflow:

1. checks that the tag matches the `package.json` version;
2. runs `npm ci`, `npm test`, `npm run validate` and `npm run build`;
3. publishes with `npm publish --provenance` through npm **trusted publishing** (OIDC,
   no npm token);
4. creates the GitHub Release. Its notes are the version's `CHANGELOG.md` section, or
   generated notes when there isn't one.

Prerelease versions (anything with a `-`) bundle the gold Insider icon; stable versions use
the regular icon (`vite.config.ts`).

## One-time setup (npmjs.com)

Trusted publishing must be configured for the package on npmjs.com: **Package settings →
Trusted publishing → GitHub Actions**, with organization or user `mohsinonxrm`,
repository `pptb-mxrm-dataverse-request-studio` and workflow filename `release.yml`
(no environment). Until it is, the publish step fails.

## Cutting a release

1. Release-prep PR into `dev`: `npm version X.Y.Z[-beta.N] --no-git-tag-version`, a dated
   `## [X.Y.Z[-beta.N]]` section in `CHANGELOG.md`, and README updates. Merge it.
2. Release PR `dev` → `main`, merged with a **merge commit** (never squash, so `dev` and
   `main` keep a shared history).
3. Tag the merge commit on `main` and push the tag:

   ```bash
   git switch main && git pull && git tag v1.4.0-beta.1 && git push origin v1.4.0-beta.1
   ```

4. Watch the **Release** workflow, then check:

   ```bash
   npm view @mohsinonxrm/pptb-dataverse-request-studio dist-tags
   ```

   and the GitHub Release (marked pre-release for `-beta.N` / `-rc.N`).

A beta stays on `beta` until it's promoted: release `X.Y.Z` (no suffix) the same way, which
publishes to `latest`.

## Branch model

Work happens on branches cut from `dev`, merged into `dev` by PR. Releases go `dev` →
`main`. Adopting the full shared model from FetchXML Studio (`release/X.Y` for betas,
`support/X.Y` for the previous minor) is tracked in #90.
