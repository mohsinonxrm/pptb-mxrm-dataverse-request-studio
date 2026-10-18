# Dataverse Request Studio for Power Platform ToolBox

A metadata-driven studio for composing, previewing, and executing Microsoft
Dataverse Web API requests — covering OData reads, record writes, table
associations, bound and unbound action/function invocation, on-demand workflows,
and file / image / attachment operations.

Built for [Power Platform ToolBox](https://www.powerplatformtoolbox.com/) with
React 18, Fluent UI v9, TypeScript, and live Dataverse metadata integration.

[![npm](https://img.shields.io/npm/v/@mohsinonxrm/pptb-dataverse-request-studio?label=npm)](https://www.npmjs.com/package/@mohsinonxrm/pptb-dataverse-request-studio)
![React](https://img.shields.io/badge/React-18-61DAFB?logo=react)
![TypeScript](https://img.shields.io/badge/TypeScript-5.9-3178C6?logo=typescript)
![Fluent UI](https://img.shields.io/badge/Fluent%20UI-v9-0078D4?logo=microsoft)
![PPTB Types](https://img.shields.io/badge/%40pptb%2Ftypes-1.2.1-orange)
![License](https://img.shields.io/badge/license-AGPL--3.0--only-green)

---

## ✨ Features

DRS ships with **19 request modes** organized into 5 functional groups. Every
mode is metadata-driven, generates the correct OData / Web API shape, and
emits ready-to-copy code in nine formats: `fetch`, `Xrm.WebApi`,
`Xrm.WebApi.executeMultiple`, `XMLHttpRequest`, Power Automate, C#, PowerShell,
cURL, and JSON.

### 📖 Read group

- **Retrieve Multiple** — full OData query composer with `$select`, `$filter`,
  `$orderby`, `$expand`, `$top`, `$count`, `$apply` (groupby / aggregate), and
  a virtualized results grid with infinite scroll via `@odata.nextLink`.
- **Retrieve Single** — one record by primary key (GUID), with `$select` /
  `$expand`.
- **Retrieve NextLink** — paste a `@odata.nextLink` URL to continue paging
  any query.
- **Predefined Query** — execute saved-query (`?savedQuery=<id>`) and
  user-query (`?userQuery=<id>`) URLs against an entity set; layout-aware
  results grid that matches the source view's column order.

### ✏️ Write group

- **Create Record** — `POST /<set>` with metadata-driven field-set editor.
  Required-field guard with red `req` badges, `Prefer: return=representation`
  toggle, optional `$select` return list, `MSCRM.*` bypass-header family.
- **Update Record** — `PATCH /<set>(<id>)` multi-field OR
  `PUT /<set>(<id>)/<col>` single-column. Diff view (before → after) with
  per-row drill-down to switch to PUT mode.
- **Upsert Record** — full upsert / create-only / update-only / etag-guarded
  variants. By GUID OR alternate key.
- **Delete Record** — `DELETE /<set>(<id>)` with a typed-confirmation safety
  affordance + cascade-impact preview. Single-property delete via
  `DELETE /<set>(<id>)/<col>` also supported.
- **Merge Records** — `POST /Merge` for `account` / `contact` / `incident`
  with target + subordinate record cards, field-override matrix, and
  PerformParentingChecks / SuppressDuplicateDetection settings.

### 🔗 Relate group

- **Associate** — links related records via `POST /<set>(<id>)/<nav>/$ref`
  for collection-valued navs, or `PATCH @odata.bind` for single-valued.
  Cardinality detected from metadata; verb auto-flips with the picker.
  Polymorphic lookups (Customer / Owner / `regardingobjectid`) fully
  supported with target-disambiguated nav properties.
- **Disassociate** — `DELETE /$ref` for collection-valued, `PATCH @odata.bind: null`
  for single-valued. Same polymorphism story.

### ⚡ Execute group

- **Execute Action** — OOB action invocation via bound
  `POST /<set>(<id>)/Microsoft.Dynamics.CRM.<Action>` OR unbound
  `POST /<Action>`. CSDL-driven parameter form with type-aware controls:
  required-first UX, Form ↔ JSON toggle, EntityReference / EntityCollection
  / Edm primitives / Enum / Status-code OptionSet enrichment.
- **Execute Custom API** / **Execute Custom Action** — the same composer,
  with the picker filtered by message source. The source is currently
  inferred from the message name (a publisher prefix such as `new_` reads as
  custom), so custom messages are listed under Custom API.
- **Execute Function** — OData function calls with parameter aliasing
  (`?@p1=...`) vs inline literal toggle. CSDL-driven param form with the
  same type-aware controls as actions.
- **Execute Workflow** — runs on-demand workflows via
  `POST /workflows(<id>)/Microsoft.Dynamics.CRM.ExecuteWorkflow` with the
  bound record id. Live picker over activated, on-demand process workflows
  (`statecode eq 1 and type eq 1 and ondemand eq true and category eq 0`).

### 💾 Binary group

- **Manage File Data** — File-column upload / download / delete with three
  upload paths (single-request PATCH, chunked PATCH with session token,
  `InitializeFileBlocksUpload` / `UploadBlock` / `CommitFileBlocksUpload`
  messages), four download paths (single GET `/$value`, ranged GET with
  `Range`, message-based block download, `GetFileSasUrl` function), and the
  `DeleteFile` message + DELETE-on-property alternatives.
- **Manage Image Data** — Image columns with thumbnail vs full-size
  download (`?size=full`), `CanStoreFullImage` awareness, and pointers to
  the file-column path for full-size uploads.
- **Manage Attachment / Note** — `activitymimeattachment.body` and
  `annotation.documentbody`. Inline base 64 OR
  `Initialize{Attachment|Annotation}BlocksUpload` /
  `Commit{Attachment|Annotation}BlocksUpload` messages. Single-request
  GET `/$value` OR `Initialize{...}BlocksDownload` / `DownloadBlock` for
  download.

> The binary modes compose the request pipeline and its code, but the PPTB
> Dataverse API doesn't expose these chunked operations, so **Execute returns
> `501 Not Implemented`** in this group. Run the generated code from the Code
> tab instead.

### 🔍 Cross-cutting capabilities

- **Live metadata** — Tables, columns, relationships, alternate keys, action
  / function CSDL all fetched lazily from the host's `dataverseAPI`. Per-org
  TTL cache with in-flight promise dedup. Type-specific column properties
  (Targets, OptionSet, MaxLength, MinValue/MaxValue, Format) loaded on-demand
  by the editor that needs them.
- **Antipattern advisories** — Surfaces real-time guidance per the
  [Microsoft query-antipatterns doc](https://learn.microsoft.com/en-us/power-apps/developer/data-platform/query-antipatterns):
  `FilteringOnCalculatedColumns`, `LargeAmountOfLogicalAttributes`, large-text
  `contains`, leading-wildcard rewrites — all flagged inline AND aggregated
  into an Advisory drawer in the URL bar.
- **Polymorphic lookup support** — Customer, Owner, and multi-target lookups
  (`regardingobjectid`) emit the correct target-disambiguated
  `<navName>@odata.bind` binding shape automatically.
- **Code generators** — Every mode emits the same request as `fetch`,
  `Xrm.WebApi` (CRUD helpers or `Xrm.WebApi.online.execute`),
  `Xrm.WebApi.online.executeMultiple`, `XMLHttpRequest`, C# `HttpClient`,
  PowerShell `Invoke-RestMethod`, cURL, and a JSON description, plus a
  field-by-field Power Automate (Dataverse connector) sheet. Multi-request
  pipelines (chunked uploads, message-based binary ops) render as full
  ordered sequences.
- **Saved-request library** — Per-org localStorage persistence with
  auto-suggested names (`account · 3 sel · 2 flt · 2026-05-23`),
  rename / delete / overwrite, dirty-state detection, and cross-tab sync.
- **PPTB host integration** — `isEmbedded()`-aware UI: theme sync,
  environment-aware URLs, and host-stripped header advisories. PPTB's
  Dataverse API doesn't send custom request headers, so Execute is blocked
  when a request depends on a precondition (`If-Match` / `If-None-Match`) or
  `MSCRM.*` bypass header — copy it from the Code tab to run it elsewhere.

---

## 🖼️ Interface overview

```
┌─────────────────────────────────────────────────────────────────────────┐
│  [Request Type ▼] [Group · Method]               [⚙ Settings] [Theme]   │
├──────────────────────────┬──────────────────────────────────────────────┤
│                          │  [Builder]  [Code]  [Results]                 │
│   Sidebar                ├──────────────────────────────────────────────┤
│                          │                                               │
│   📋 Target              │   Active editor pane                          │
│      account             │                                               │
│   📊 $select       3 sel │   ┌────────────────────────────────────────┐ │
│   🔍 $filter       2 flt │   │ (clause editor — varies per mode)      │ │
│   🔁 $orderby      1 sort│   │                                        │ │
│   🔗 $expand       1 exp │   │  • Field-set editor (writes)           │ │
│   📃 $top                │   │  • Filter tree (reads)                 │ │
│   ⚙ Prefer               │   │  • Action param form (executes)        │ │
│   📥 Headers       1 act │   │  • Binary pipeline (file/image)        │ │
│                          │   └────────────────────────────────────────┘ │
│   ⏱ Recent runs          │                                               │
│      200 · 142ms · 32B   │                                               │
│      200 · 89ms          │                                               │
│                          │                                               │
├──────────────────────────┴──────────────────────────────────────────────┤
│  GET /accounts?$select=name,…                  0.4 KB  ⚠ 2 advisories  │
│                                                       [▶ Execute]      │
└─────────────────────────────────────────────────────────────────────────┘
```

---

## 🚀 Getting started

### Prerequisites

- [Power Platform ToolBox](https://www.powerplatformtoolbox.com/) **1.2.0 or
  later** (the tool declares `minAPI: 1.2.0`)
- A connection to a Dataverse environment
- For local development: **Node.js 20.19+ or 22.12+** (Vite 7) and **npm**

### Installation

Dataverse Request Studio is available as a tool in Power Platform ToolBox.
Install it from the tool gallery or load it as a custom tool.

---

## 🛠️ Development

### Setup

```bash
git clone https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio.git
cd pptb-mxrm-dataverse-request-studio
npm install
```

### Common scripts

```bash
# Vite dev server (standalone, no PPTB host)
npm run dev

# TypeScript-only check (fast)
npm run typecheck

# Production build
npm run build

# Preview the production build locally
npm run preview

# Validate manifest against PPTB registry rules
npm run validate

# Pre-publish: validate → build
npm run finalize-package
```

### Standalone vs PPTB-embedded

DRS runs in two modes:

- **Standalone** (`npm run dev`) — opens in a regular browser tab. Useful
  for UI iteration only: without the host bridge no live metadata loads, and
  Execute returns an error result (HTTP 500, "PPTB host method … is not
  available — running standalone").
- **Embedded in PPTB** — loads inside the PPTB iframe. The host injects
  `window.dataverseAPI` (typed) and `window.toolboxAPI`. Execute runs
  against the active Dataverse connection.

See [docs/CONTRIBUTING.md](docs/CONTRIBUTING.md) for branch conventions,
commit style, code style, and PR guidance.

---

## 📦 Tech stack

| Technology | Version | Purpose |
|---|---|---|
| **React** | 18.3 | UI framework |
| **TypeScript** | 5.9 | Strict mode, discriminated unions |
| **Vite** | 7 | Build tooling + HMR |
| **Fluent UI v9** | 9.73 | Microsoft design system (Combobox, DataGrid, Tree, Drawer, Tabs, etc.) |
| **Monaco Editor** | 0.54 | JSON / code editor for body authoring and Code tab |
| **@fluentui-contrib/react-data-grid-react-window** | 1.4 | Virtualized results grid |
| **@dnd-kit** | 6.3 | Drag-reordering for filter / orderby / select clauses |
| **@pptb/types** | 1.2.1 | Power Platform ToolBox host API types |

---

## 📁 Project layout

```
src/
  App.tsx                    — root, ThemedApp, Frame, mode router
  main.tsx                   — Vite entry point

  modes/                     — request-mode components (19 modes)
    RetrieveMultipleMode.tsx
    RetrieveSingleMode.tsx
    RetrieveNextLinkMode.tsx
    PredefinedQueryMode.tsx
    CreateMode.tsx / UpdateMode.tsx / UpsertMode.tsx / DeleteMode.tsx
    MergeMode.tsx
    AssociateMode.tsx / DisassociateMode.tsx
    ExecuteActionMode.tsx    — Execute Action / Custom API / Custom Action
    ExecuteFunctionMode.tsx / ExecuteWorkflowMode.tsx
    ManageFileMode.tsx / ManageImageMode.tsx / ManageAttachmentMode.tsx
    StubMode.tsx             — fallback for an unknown mode id

  editors/                   — clause editors shared across modes
    filter/                  — FilterEditor + value inputs + operator catalog
    expand/                  — $expand tree navigation
    FieldSetEditor.tsx       — metadata-driven write-side form
    SelectEditor.tsx, OrderbyEditor.tsx, ExpandEditor.tsx, ApplyEditor.tsx
    ActionParamForm.tsx      — CSDL-driven action / function params
    BinarySourceCard.tsx, BinaryPipelineCard.tsx, BinaryColumnPickers.tsx
    HeadersEditor.tsx, BypassEditor.tsx, PreconditionEditor.tsx
    ActionPicker.tsx, NavPropertyPicker.tsx, PasteODataDialog.tsx
    ...

  engine/                    — URL / body construction + code generation
    urlBuilder.ts            — central URL builder for every mode
    executeBuilders.ts       — action / function / workflow body builders
    binaryBuilders.ts        — file / image / attachment pipelines
    codeGenerators.ts        — fetch, Xrm.WebApi (+ executeMultiple), XHR,
                               Power Automate, C#, PowerShell, cURL, JSON
    runtime.ts               — single dispatch point for execute
    dataverseExecutor.ts     — live executor over the host's dataverseAPI
    odataParser.ts           — round-trips OData URL strings ↔ state
    antipatterns.tsx, bypassAdvisories.tsx, hostHeaderAdvisories.tsx
                             — advisory detectors

  host/                      — PPTB bridge + metadata loaders
    pptbBridge.ts            — typed wrapper over window.dataverseAPI
    metadataProvider.ts      — entity / attribute / relationship loaders
    csdlProvider.ts          — $metadata XML parsing + action / function index
    cache.ts                 — singleton TTL cache with in-flight dedup
    useLiveMetadata.ts, useColumnDetail.ts, useScopedEntities.ts, …

  primitives/                — small reusable UI bits
    ClauseTreeItem.tsx, SectionHeader.tsx, ModeCard.tsx, RequestTypePicker.tsx
    StatusPill.tsx, MethodPill.tsx, AdvisoryDrawer.tsx, advisories.ts
    NavPathColumnPicker.tsx  — N:1 drill-down column picker for filter / apply
    RecordPicker.tsx, KvGrid.tsx, Sortable.tsx, SegmentedToggle.tsx, …

  shell/                     — header / sidebar / URL bar / tabs / drawers
    FrameHeader.tsx, ModeShell.tsx, Sidebar.tsx
    UrlBar.tsx, UrlTokens.tsx, MainTabs.tsx, SettingsDrawer.tsx

  state/                     — per-mode state shapes + persistence
    readState.ts, writeState.ts, relateState.ts, executeState.ts, binaryState.ts
    savedRequests.ts         — localStorage persistence (per-org scope)
    SaveContext.tsx          — cross-mode save/load context
    validateRequest.tsx      — pre-execute request validation
    displaySettings.ts       — display preferences

  registry/                  — request-type catalogue
    requestTypes.ts          — 19 request types grouped into 5 groups

  views/                     — output panes
    CodeView.tsx             — Code tab: one sub-tab per code format
    PowerAutomatePane.tsx    — Power Automate field-by-field sheet
    ResultsView.tsx          — tabs (Detail/Grid/JSON/Headers) + error state
    results/                 — grid + cell renderers + flattening
    detail/                  — record detail card

  theme/                     — Fluent theme + studio-owned brand vars
```

---

## 📋 Web API coverage

✅ = compose + Execute in PPTB · 📝 = compose + Code tab only (Execute is
blocked or not available through the PPTB Dataverse API)

| Operation | Status |
|---|---|
| Retrieve Multiple (full OData) | ✅ |
| Retrieve Single (by GUID) | ✅ |
| `@odata.nextLink` paging | ✅ |
| Predefined / user queries | ✅ |
| `$apply` (groupby / aggregate) | ✅ |
| Lambda filter (`any` / `all`) | ✅ |
| `$expand` (single + collection, nested) | ✅ |
| Polymorphic lookups (Customer, Owner, regardingobjectid) | ✅ |
| Create / Update / Upsert / Delete | ✅ |
| `Prefer: return=representation` | ✅ |
| Optimistic concurrency (`If-Match` ETag) | 📝 |
| `MSCRM.*` bypass family (custom plug-ins, flows, async) | 📝 |
| Single-property PUT / DELETE | ✅ |
| Alternate-key upsert | ✅ |
| Associate / Disassociate (collection + single-valued) | ✅ |
| Merge (account / contact / incident) | ✅ |
| Execute Action (OOB / Custom API / Custom Action) | ✅ |
| Execute Function | ✅ |
| Execute Workflow | ✅ |
| File column (upload / download / delete) | 📝 |
| Image column (thumbnail / full-size) | 📝 |
| Attachment / Note (inline + message-based) | 📝 |
| `$batch` envelope composer | ⏳ (see roadmap) |
| `CreateMultiple` / `UpdateMultiple` / `DeleteMultiple` | ⏳ |
| `BulkDelete` async job | ⏳ |
| `ExecuteMultiple` envelope | ⏳ |

---

## 🤝 Contributing

See [docs/CONTRIBUTING.md](docs/CONTRIBUTING.md) for branch conventions,
commit style, code style, repo layout, and PR guidance.

---

## 📋 Roadmap

- `$batch` envelope composer mode
- Bulk-write modes (`CreateMultiple`, `UpdateMultiple`, `DeleteMultiple`)
- `BulkDelete` async-job composer
- `ExecuteMultiple` request envelope
- Recent-runs click-to-reload
- WriteResultCard "Retrieve this record" cross-mode jump
- EntitySpecific param — reference vs new-instance toggle
- Workflow input parameters (parse XAML `InArgument`)
- Predefined Query — apply to related collection
- CSV / JSON export from results grid
- Column resize with persisted widths
- Smoke-test suite covering URL builders + code generators

---

## 📄 License

Licensed under the **GNU Affero General Public License v3.0** (AGPL-3.0-only).
See [LICENSE](LICENSE) for details.

---

## 🙏 Acknowledgments

- [Dataverse REST Builder](https://github.com/GuidoPreite/DRB) - Original inspiration.
- [Power Platform ToolBox](https://www.powerplatformtoolbox.com/) — host
  platform for the desktop / web app where DRS runs.
- [FetchXML Studio](https://github.com/mohsinonxrm/pptb-mxrm-fetchxml-studio) — sibling tool that pioneered many of the patterns
  reused here (metadata cache, scope settings, results grid).
- [Fluent UI](https://react.fluentui.dev/) — UI component library.
- [Monaco Editor](https://microsoft.github.io/monaco-editor/) — VS Code's editor core.

---

Built with ❤️ for the Power Platform community.
