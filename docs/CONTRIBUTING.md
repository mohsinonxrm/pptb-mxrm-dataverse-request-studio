# Contributing to Dataverse Request Studio

Thanks for your interest in contributing. This document covers the basics —
how the repo is laid out, the branch model, commit conventions, and how to
run the studio locally.

For bug reports and feature requests, please use the GitHub issue templates.

---

## Quick start

### Prerequisites

- **Node.js** 22.22.2+, 24.15+ or 26+ (`engines` in `package.json`)
- **npm** (bundled with Node)
- A modern browser (Edge, Chrome, Firefox, or Safari)

### Set up

```bash
git clone https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio.git
cd pptb-mxrm-dataverse-request-studio
npm install
```

### Common scripts

| Command | What it does |
|---|---|
| `npm run dev` | Vite dev server on `localhost:5173` (standalone, no PPTB host) |
| `npm run typecheck` | `tsc -b --noEmit` — fast type-only check |
| `npm test` / `npm run test:watch` | Vitest unit tests, once / in watch mode |
| `npm run build` | Full production build into `dist/` |
| `npm run preview` | Serve the production build locally |
| `npm run format` / `npm run format:check` | Prettier write / check over the repo |
| `npm run validate` | Run `pptb-validate` against the manifest |
| `npm run finalize-package` | `validate` → `build` (pre-release) |
| `npm run build:analyze` | Production build plus a bundle-size report |

### Standalone vs PPTB-embedded

The studio runs in two modes:

- **Standalone** (`npm run dev`) — opens in a regular browser tab.
  Useful for UI iteration. Execute will not actually call Dataverse because
  the host bridge is not present; you can still build requests, preview
  URLs, and use the Code tab.
- **Embedded in PPTB** — the studio loads inside the Power Platform ToolBox
  iframe. The host injects a `window.dataverseAPI` typed surface and a
  `window.toolboxAPI` for filesystem / settings. Execute runs against the
  active Dataverse connection.

The bridge layer in `src/host/` abstracts the difference. `isEmbedded()`
returns true only in the iframe.

---

## Branch model

- **`dev`** — default working branch. All feature work, fixes, and refactors
  target `dev`. PRs into `dev` should be small, focused, and pass the CI
  checks (see [Pull requests](#pull-requests)).
- **`main`** — protected. Only updated at release time by a PR from `dev`,
  merged with a merge commit. A release is a `vX.Y.Z` (or `vX.Y.Z-beta.N`)
  tag on that merge commit: the Release workflow publishes it to npm and
  creates the GitHub Release. Nobody publishes by hand; see
  [RELEASING.md](../RELEASING.md).

### Workflow

```bash
git switch dev
git pull
git switch -c feat/<short-description>
# … work …
git push -u origin feat/<short-description>
# Open PR → dev
```

---

## Commit conventions

We loosely follow [Conventional Commits](https://www.conventionalcommits.org/):

```
<type>(<optional scope>): <short summary>

<optional body>
```

Common types:

- `feat:` new user-facing feature
- `fix:` bug fix
- `refactor:` code change that neither adds a feature nor fixes a bug
- `chore:` housekeeping (deps, build config, etc.)
- `docs:` documentation only
- `style:` formatting only (no behavior change)
- `perf:` performance improvement

Examples:

```
feat(retrieve-multiple): infinite scroll via @odata.nextLink
fix(filter): wildcard not stripped for endswith
refactor(savedRequests): scope storage key per Dataverse org
chore(deps): bump @fluentui/react-components to 9.74
```

Keep subject lines under ~72 characters. Use the body for the *why* when
the *what* is obvious from the diff.

---

## Code style

- **TypeScript strict mode** — avoid `any`; where it's unavoidable, add a
  comment explaining why. (The repo has no ESLint setup, so existing
  `eslint-disable` comments are not enforced by anything.)
- **Prettier** formats everything — run `npm run format` before pushing;
  `npm run format:check` must pass.
- **Fluent UI v9** primitives — prefer `@fluentui/react-components` over
  hand-rolled DOM. Use `tokens.*` for colors; studio-owned CSS variables
  in `index.html` for brand / method palettes.
- **Discriminated unions** for domain-shaped data (`ColumnMeta`,
  `FileUploadMethod`, `Advisory`, etc.). Exhaustive switch with `never`
  fallthrough.
- **State shape** lives in `src/state/`, one file per mode group.
  Modes never share state via React context — only via the explicit
  `usePublishSaveContext` channel.
- **Comments** explain *why*, not *what*. Cite public MS Learn docs when
  the behavior comes from an external spec; never cite internal planning
  artifacts.

---

## Repo layout

```
src/
  App.tsx             — root + ThemedApp + Frame
  main.tsx            — Vite entry point
  editors/            — composition pieces shared across modes
  engine/             — URL builders, code generators, runtime dispatch
  host/               — PPTB bridge + live metadata loaders + cache
  mock/               — type definitions + live registry (no fixtures)
  modes/              — request-mode components (23 modes; Execute Action /
                        Custom API / Custom Action share one component)
  primitives/         — small UI building blocks (pickers, pills, advisories)
  registry/           — request-type catalogue + group definitions
  shell/              — header / sidebar / URL bar / main tabs / settings
  state/              — per-mode state shapes + saved-request persistence
  theme/              — Fluent theme + studio-owned tokens
  views/              — results pane + code pane
docs/
  CONTRIBUTING.md     — this file
icons/                — tool icons (bundled into dist/icons at build time)
.github/              — issue / PR templates
```

---

## Pull requests

A good PR:

- Targets `dev`
- Passes `npm run typecheck`, `npm run format:check`, `npm test` AND
  `npm run build` locally. CI (`.github/workflows/ci.yml`) runs these plus
  `npm run validate` on every PR, and GitHub code scanning (CodeQL) runs too
- Touches one logical concern (split mechanical refactors into their own PR)
- Includes screenshots for any UI change (light AND dark mode)
- Updates `CHANGELOG.md` under `## [Unreleased]` if user-facing
- Has a self-review pass before requesting reviewers

The maintainer merges approved PRs with a merge commit (no squash).

---

## Reporting bugs / requesting features

Open a GitHub issue using the appropriate template:

- **Bug report** — include repro steps, expected vs actual, environment
  (browser, OS, Dataverse env URL if relevant), and screenshots
- **Dataverse Web API request bug** — the URL / body / headers / code DRS
  generated is wrong, or Dataverse rejects it
- **Performance issue** — slow UI, metadata loading, or execution
- **Feature request** — motivation, proposed UX, alternatives considered
- **Documentation** — missing or incorrect docs
- **Question** — for usage questions before they're a confirmed bug

Security issues go through GitHub Security Advisories — see
[SECURITY.md](../SECURITY.md).

---

## License

This project is licensed under **AGPL-3.0-only**. By contributing you agree
your contributions are licensed under the same terms.
