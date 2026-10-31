# Dataverse Request Studio for Power Platform ToolBox

A metadata-driven studio for composing, previewing, and executing Microsoft
Dataverse Web API requests — covering OData reads, record writes, bulk writes
and bulk-delete jobs, table associations, bound and unbound action/function
invocation, on-demand workflows, `$batch` / ExecuteMultiple envelopes, and
file / image / attachment operations.

Built for [Power Platform ToolBox](https://www.powerplatformtoolbox.com/) with
React 18, Fluent UI v9, TypeScript, and live Dataverse metadata integration.

[![npm](https://img.shields.io/npm/v/@mohsinonxrm/pptb-dataverse-request-studio?label=npm)](https://www.npmjs.com/package/@mohsinonxrm/pptb-dataverse-request-studio)
[![npm beta](https://img.shields.io/npm/v/@mohsinonxrm/pptb-dataverse-request-studio/beta?label=beta&color=orange)](https://www.npmjs.com/package/@mohsinonxrm/pptb-dataverse-request-studio?activeTab=versions)
![React](https://img.shields.io/badge/React-18-61DAFB?logo=react)
![TypeScript](https://img.shields.io/badge/TypeScript-5.9-3178C6?logo=typescript)
![Fluent UI](https://img.shields.io/badge/Fluent%20UI-v9-0078D4?logo=microsoft)
![PPTB Types](https://img.shields.io/badge/%40pptb%2Ftypes-1.2.7-orange)
![License](https://img.shields.io/badge/license-AGPL--3.0--only-green)

---

## ✨ Features

DRS ships with **23 request modes** in 7 groups. Every mode is
metadata-driven and generates the correct OData / Web API shape. The
single-request modes emit ready-to-copy code in eleven formats: `fetch`,
`Xrm.WebApi`, `Xrm.WebApi.executeMultiple`, `XMLHttpRequest`, Power Automate,
C#, Python (`requests`), PowerShell, cURL, JSON, and an OpenAPI 3 document.
`$batch` has its own HTTP / fetch / Python / PPTB outputs, and ExecuteMultiple
has C# / HTTP / Python / PPTB.

### 📖 Read group

- **Retrieve Multiple** — full OData query composer with `$select`, `$filter`,
  `$orderby`, `$expand` (nested, with inner `$filter` / `$orderby` / `$top`),
  `$top`, `$count`, `$apply` (groupby / aggregate), and a virtualized results
  grid with infinite scroll via `@odata.nextLink`, **Retrieve all** (pauses
  every 50,000 rows with **Continue**), CSV / JSON export, per-column filters,
  row search, density and persisted column widths.
  Address shortcut chips in the `$filter` builder; paste any OData URL to load
  it into the builder.
- **Retrieve Single** — one record by primary key (GUID), with `$select` /
  `$expand`.
- **Retrieve NextLink** — paste a `@odata.nextLink` URL to continue paging
  any query.
- **Predefined Query** — execute saved-query (`?savedQuery=<id>`) and
  user-query (`?userQuery=<id>`) URLs against an entity set, or over a row's
  related records; layout-aware results grid that matches the source view's
  column order.

### ✏️ Write group

- **Create Record** — `POST /<set>` with a metadata-driven field-set editor or
  a raw JSON body. Required-field guard, `Prefer: return=representation`,
  optional `$select` return list, duplicate detection and the `MSCRM.*`
  bypass-header family. **Retrieve this record** from the result card.
- **Update Record** — `PATCH /<set>(<id>)` multi-field OR
  `PUT /<set>(<id>)/<col>` single-column (Code tab only), with a before →
  after diff.
  **Require existing record** (`If-Match: *`) is the default, so a wrong GUID
  returns 404 instead of creating a record.
- **Upsert Record** — full upsert / create-only / update-only / ETag-guarded
  variants, by GUID or alternate key (alternate key: Code tab only).
- **Delete Record** — `DELETE /<set>(<id>)` with a typed-confirmation safety
  affordance + cascade-impact preview. Single-property delete via
  `DELETE /<set>(<id>)/<col>` is composed for the Code tab.
- **Merge Records** — `POST /Merge` for `account` / `contact` / `lead` /
  `incident` with target + subordinate record cards, field-override matrix,
  and PerformParentingChecks / SuppressDuplicateDetection settings.

### 📦 Bulk group

- **Bulk Write** — `CreateMultiple` / `UpdateMultiple` / `UpsertMultiple` /
  `DeleteMultiple` over a grid of rows, with a per-table support check.
- **Bulk Delete Job** — a `BulkDelete` system job built from a query, with a
  scheduled or immediate start and live job-status polling.

### 🔗 Relate group

- **Associate** — links related records via `POST /<set>(<id>)/<nav>/$ref`
  for collection-valued navs (one request per target, already-linked targets
  skipped), or `PATCH @odata.bind` for single-valued. Cardinality detected
  from metadata; verb auto-flips with the picker. Polymorphic lookups
  (Customer / Owner / `regardingobjectid`) fully supported.
- **Disassociate** — `DELETE /$ref` for collection-valued, `PATCH @odata.bind: null`
  for single-valued. Same polymorphism story.
- Single-valued PATCHes carry `If-Match: *`, so a wrong source id can't create
  a record. Multi-target results are reported per target.

### ⚡ Execute group

- **Execute Action** — OOB action invocation via bound
  `POST /<set>(<id>)/Microsoft.Dynamics.CRM.<Action>`, collection-bound, OR
  unbound `POST /<Action>`. CSDL-driven parameter form with type-aware
  controls: required-first UX, Form ↔ JSON toggle, EntityReference (existing
  record) or EntitySpecific (existing record or new instance) /
  EntityCollection / Edm primitives / collections / Enum / OptionSet
  enrichment.
- **Execute Custom API** / **Execute Custom Action** — the same composer,
  listing every Custom API and every custom process action, Microsoft's own
  included. Classification comes from Dataverse data (`customapi`,
  `workflow` + `sdkmessage`); the name is only a fallback when those lookups
  fail.
- **Execute Function** — OData function calls with parameter aliasing
  (`?@p1=...`) vs inline literals, valid OData literals and escaping, run from
  the built URL.
- **Execute Workflow** — runs on-demand workflows via
  `POST /workflows(<id>)/Microsoft.Dynamics.CRM.ExecuteWorkflow` with the
  bound record id, showing the workflow's declared input arguments.

### 🧺 Batch group

- **`$batch` Request** — compose a `$batch` from groups of requests: change
  sets (atomic, with `$n` Content-ID references) or independent requests, with
  continue-on-error. Add the last request of a read, write, relate or execute
  mode; per-part results with
  status and body. A batch that mixes change sets and independent requests
  is sent as consecutive ToolBox calls. Needs PPTB 1.2.7+.
- **ExecuteMultiple** — the ExecuteMultiple pattern (independent requests,
  continue-on-error, return responses) on the `$batch` engine, with SDK
  `ExecuteMultipleRequest` (C#), raw `$batch` and PPTB code. Needs PPTB
  1.2.7+.

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
  / function CSDL all fetched lazily from the host's `dataverseAPI`, cached per
  org with in-flight promise dedup. Switching the PPTB connection resets the
  prepared request and reloads metadata for the new environment.
- **Request headers through PPTB 1.2.7+** — Create / Update / Upsert /
  Delete / Merge / Bulk Write / Execute Action / Execute Workflow headers
  (preconditions, `MSCRM.*` bypass, duplicate detection, impersonation) are
  sent as additional headers after you approve them once in ToolBox. The
  read modes, Execute Function, Associate / Disassociate and Bulk Delete Job send an
  allow-list: impersonation (`MSCRMCallerID`, `CallerObjectId`) everywhere,
  plus `If-None-Match`, `Consistency` and the page size on reads, and the
  bypass headers on Associate / Disassociate; the Headers pane marks the
  rest as Code tab only. On older hosts, requests that depend on a
  precondition, bypass or impersonation header are blocked with an
  explanation. `Authorization` / `Cookie` headers are never sent, saved or
  exported.
- **Antipattern advisories** — Surfaces real-time guidance per the
  [Microsoft query-antipatterns doc](https://learn.microsoft.com/en-us/power-apps/developer/data-platform/query-antipatterns):
  `FilteringOnCalculatedColumns`, `LargeAmountOfLogicalAttributes`, large-text
  `contains`, leading-wildcard rewrites — all flagged inline AND aggregated
  into an Advisory drawer in the URL bar.
- **Polymorphic lookup support** — Customer, Owner, and multi-target lookups
  (`regardingobjectid`) emit the correct target-disambiguated
  `<navName>@odata.bind` binding shape automatically.
- **Code generators** — Every single-request mode emits the same request as `fetch`,
  `Xrm.WebApi` (CRUD helpers or `Xrm.WebApi.online.execute`),
  `Xrm.WebApi.online.executeMultiple`, `XMLHttpRequest`, C# `HttpClient`,
  PowerShell `Invoke-RestMethod`, cURL, a JSON description and an OpenAPI 3
  document, plus a field-by-field Power Automate (Dataverse connector) sheet.
  Multi-request pipelines render as full ordered sequences.
- **Saved-request library** — per-connection, versioned, up to 200 requests
  across all modes, with auto-suggested names, rename / delete / overwrite and
  dirty-state detection. Export / import as JSON, export as a Postman /
  Insomnia / Bruno collection, and copy or move requests from other
  connections. A library that can't be read is never overwritten.
- **Recent runs** — click a run to restore it into its mode. Runs are kept
  across restarts, per environment (the last 200, without response bodies or
  credential / impersonation headers), and **Show all** searches them. Turn
  the history off or clear it in Settings. Each mode keeps its unsaved request
  when you switch modes.
- **Results** — CSV export (formulas neutralised for spreadsheets), JSON
  export, Copy as TSV, and response-size hints.
- **Accessibility** — labelled inputs, readable badges and live-region
  announcements; render errors are contained per pane.

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
  later** (the tool declares `minAPI: 1.2.0`). **1.2.7 or later** is
  recommended: it sends request headers (impersonation, preconditions,
  bypass) and the page size, and is required for
  the `$batch` and ExecuteMultiple modes. DRS detects what the host can do
  at runtime, so `minAPI` stays at 1.2.0 rather than locking out older
  hosts (#273).
- A connection to a Dataverse environment
- For local development: **Node.js 22.22.2+, 24.15+ or 26+** and **npm** (see
  `engines` in `package.json`)

### Installation

Dataverse Request Studio is available as a tool in Power Platform ToolBox.
Install it from the tool gallery or load it as a custom tool. Stable releases
are published to npm under the `latest` tag. Prereleases of the next version
are published under the `beta` tag (for example `1.5.0-beta.1`) and use the
gold Insider icon.

### Choose the entity scope

The entity pickers list the entities in the scope shown on the **scope button**
in the header, after the request type. On first run it reads **Choose scope**
and a tip points at it. Open it to pick where entities come from:

- **Publisher → solution** (default): one or more publishers, then their
  solutions;
- **Solution only**: one or more solutions;
- **All entities**: everything in the environment.

**Advanced Find only** is there too. Changes apply right away, and the scope is
remembered per environment.

---

## ✏️ What this tool changes

Dataverse Request Studio only reads from Dataverse until you send a request
from one of the write, relate, execute, bulk or batch modes. Every request runs
as the connected user (or the user an impersonation header names), so Dataverse
applies that user's security roles.

**Before a write is sent, it asks.** The confirmation names the target: the
environment (connection name, URL and type, with Production highlighted), the
operation, the table and the scope. The header shows the connection and its
type at all times.

| Request | When it asks |
|---|---|
| **Merge**, **Execute Action / Custom API / Custom Action**, **Execute Workflow**, **Bulk Write**, `$batch` / **ExecuteMultiple** with any write, **Associate / Disassociate** with more than one target | Always |
| **Create**, **Update**, **Upsert**, single-target **Associate / Disassociate** | On a Production connection, with **Production safe mode** on (the default; Settings → Safety) |
| **Delete** | Always for a whole row: type the record's name and acknowledge, in the mode itself. On a Production connection with safe mode on, the confirmation above as well |
| **Bulk Delete Job** | Always: the confirmation names the environment, counts the matching rows, and you type the table name |
| Reads (Retrieve, Predefined Query, Execute Function, downloads) | Never |

On a Production connection with safe mode on, and for more than 100 records (or
more than 100 writes in a batch) on any connection, you type the table name (or
the action name, or `batch` / `ExecuteMultiple`) before the confirm button
enables. Execute and Ctrl+Enter go through the same confirmation. If the
connection changes while the confirmation is open, nothing is sent.

What each request does in Dataverse is exactly what the URL bar and the Code
tab show: one Web API request (or one per target for Associate / Disassociate,
or one `$batch` per group). The binary modes (file, image, attachment) don't
send anything from inside ToolBox: Execute returns 501, and you copy the request
from the Code tab.

**Stored on your machine, not in Dataverse:**

- **Saved requests** — in the browser storage of the ToolBox tool window, one
  library per environment (`drs:saved-requests:v1:<host>`). If a library can't
  be read, it's copied to a backup key instead of being overwritten. Export and
  Import write and read JSON files you pick.
- **Settings** (logical names, value display, entity scope, Advanced Find,
  density, column widths, Production safe mode) — in ToolBox's settings store
  for this tool.
- **Exports** (CSV, JSON, Postman, Insomnia, Bruno, OpenAPI) — only to files you
  choose. Credential headers (`Authorization`, cookies) are never saved or
  exported.

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

# Unit tests (Vitest)
npm test

# Production build
npm run build

# Bundle size report
npm run build:analyze

# Preview the production build locally
npm run preview

# Validate manifest against PPTB registry rules
npm run validate

# Lint (ESLint, no warnings allowed)
npm run lint
```

### Standalone vs PPTB-embedded

DRS runs in two modes:

- **Standalone** (`npm run dev`) — opens in a regular browser tab. Useful
  for UI iteration only: without the host bridge no live metadata loads, and
  Execute returns `501 Not Implemented` ("Running standalone: Execute needs
  the PPTB host"); the Bulk modes report it as an error result.
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
| **@pptb/types** | 1.2.7 | Power Platform ToolBox host API types |
| **Vitest** | 5 | Unit tests (request engine, code generators, storage) |

---

## 📁 Project layout

```
src/
  App.tsx                    — root, ThemedApp, Frame, mode router
  main.tsx                   — Vite entry point

  modes/                     — request-mode components (23 modes)
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
    BulkWriteMode.tsx / BulkDeleteMode.tsx
    BatchMode.tsx + batch/   — $batch Request and ExecuteMultiple
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
                               Power Automate, C#, Python, PowerShell, cURL, JSON
    openApiGenerator.ts      — OpenAPI 3 document for the current request
    collectionFormats.ts, savedToCollection.ts
                             — Postman / Insomnia / Bruno collection export
    bulkWrite*.ts, bulkDelete*.ts, batch/
                             — Bulk Write, Bulk Delete Job, $batch engine
    sensitiveHeaders.ts      — credential / impersonation header policy
    runtime.ts               — single dispatch point for execute
    dataverseExecutor.ts     — live executor over the host's dataverseAPI
    odataParser.ts           — round-trips OData URL strings ↔ state
    antipatterns.tsx, bypassAdvisories.tsx, hostHeaderAdvisories.tsx
                             — advisory detectors

  host/                      — PPTB bridge + metadata loaders
    pptbBridge.ts            — typed wrapper over window.dataverseAPI
    metadataProvider.ts      — entity / attribute / relationship loaders
    csdlProvider.ts          — $metadata XML parsing + action / function index
    cache.ts                 — per-org metadata caches (TTL + in-flight dedup)
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
    useSavedRequests.tsx     — shared saved-request store
    modeDrafts.ts            — unsaved request per mode (this session)
    SaveContext.tsx          — cross-mode save/load context
    validateRequest.tsx      — pre-execute request validation
    displaySettings.ts       — display preferences

  registry/                  — request-type catalogue
    requestTypes.ts          — 23 request types grouped into 7 groups

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

✅ = compose + Execute in PPTB · ✅ 1.2.7+ = Execute needs PPTB 1.2.7 or
later · 📝 = compose + Code tab only (not available through the PPTB
Dataverse API)

| Operation | Status |
|---|---|
| Retrieve Multiple (full OData) | ✅ |
| Retrieve Single (by GUID) | ✅ |
| `@odata.nextLink` paging, Retrieve all | ✅ |
| `Prefer: odata.maxpagesize` (page size) | ✅ 1.2.7+ (older ToolBox: Execute keeps `$top` instead) |
| Predefined / user queries (incl. over related records) | ✅ |
| `$apply` (groupby / aggregate) | ✅ |
| Lambda filter (`any` / `all`) | ✅ |
| `$expand` (single + collection, nested) | ✅ |
| Polymorphic lookups (Customer, Owner, regardingobjectid) | ✅ |
| Create / Update / Upsert / Delete | ✅ |
| `Prefer: return=representation` | ✅ (ToolBox always sends it on writes) |
| Optimistic concurrency (`If-Match` / `If-None-Match`) | ✅ 1.2.7+ |
| `MSCRM.*` bypass family (custom plug-ins, flows, async) | ✅ 1.2.7+ |
| Impersonation (`MSCRMCallerID` / `CallerObjectId`) on every mode | ✅ 1.2.7+ |
| Single-property PUT / DELETE | 📝 |
| Alternate-key upsert | 📝 |
| Associate / Disassociate (collection + single-valued) | ✅ |
| Merge (account / contact / lead / incident) | ✅ |
| Execute Action (OOB / Custom API / Custom Action) | ✅ |
| Execute Function | ✅ |
| Execute Workflow | ✅ |
| `CreateMultiple` / `UpdateMultiple` / `UpsertMultiple` / `DeleteMultiple` | ✅ |
| `BulkDelete` async job | ✅ |
| `$batch` (change sets + independent requests) | ✅ 1.2.7+ |
| `ExecuteMultiple` pattern | ✅ 1.2.7+ |
| File column (upload / download / delete) | 📝 |
| Image column (thumbnail / full-size) | 📝 |
| Attachment / Note (inline + message-based) | 📝 |

---

## 🤝 Contributing

See [docs/CONTRIBUTING.md](docs/CONTRIBUTING.md) for branch conventions,
commit style, code style, repo layout, and PR guidance.

---

## 📋 Roadmap

Planned work is tracked in the repository's
[milestones](https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/milestones)
and [issues](https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues).
Ideas under consideration include:

- Request variables (`{{var}}`) and persistent request history
- FetchXML and Dataverse Search API modes
- A metadata browser mode
- Comparing responses across runs or environments
- Importing Postman / Insomnia collections and cURL commands
- More code-generation targets (C# ServiceClient / IOrganizationService,
  Python, the PowerShell Dataverse module)

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
