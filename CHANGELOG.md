# Changelog

All notable changes to Dataverse Request Studio will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2026-05-24

Initial public release of Dataverse Request Studio — a metadata-driven studio
for composing, previewing, and executing Microsoft Dataverse Web API requests
inside Power Platform ToolBox.

### ✨ Added

#### 📖 Read group (4 modes)

- **Retrieve Multiple** — full OData composer:
  `$select` / `$filter` / `$orderby` / `$expand` / `$top` / `$skip` /
  `$count` / `$apply` (groupby + aggregate). Filter tree with AND/OR
  combinators, group nesting, NOT toggles, column-vs-column comparisons,
  lambda `any`/`all` on collection-valued navs, and the full Dataverse
  query-function family (`Microsoft.Dynamics.CRM.*`). Virtualized results
  grid with infinite scroll via `@odata.nextLink`.
- **Retrieve Single** — by primary key OR alternate key (composite alternate
  keys supported via `(col1='v1',col2='v2')` URL syntax).
- **Retrieve NextLink** — paste a server-emitted `@odata.nextLink` URL to
  continue paging any prior query.
- **Predefined Query** — execute saved-query (`?savedQuery=<id>`) and
  user-query (`?userQuery=<id>`) requests with a layout-aware results grid
  that mirrors the source view's column order from `layoutxml`.

#### ✏️ Write group (5 modes)

- **Create Record** — `POST /<set>` with metadata-driven `FieldSetEditor`,
  required-field guard, `Prefer: return=representation` toggle, return-set
  `$select`, Form ↔ JSON body-mode switcher (Monaco editor for JSON edits),
  and the `MSCRM.*` bypass-header family.
- **Update Record** — `PATCH /<set>(<id>)` multi-field OR `PUT /<set>(<id>)/<col>`
  single-column. Diff view (before → after) with per-row drill-down that
  flips the method pill from PATCH to PUT automatically.
- **Upsert Record** — full upsert / create-only / update-only / etag-guarded
  variants. By GUID OR alternate key.
- **Delete Record** — `DELETE /<set>(<id>)` with a typed-confirmation safety
  affordance + cascade-impact preview. Single-property delete via
  `DELETE /<set>(<id>)/<col>` and `PUT /<set>(<id>)/<col>` with `{"value": null}`
  also supported.
- **Merge Records** — `POST /Merge` for `account` / `contact` / `incident`
  with target + subordinate record cards, field-override matrix, and
  PerformParentingChecks / SuppressDuplicateDetection settings.

#### 🔗 Relate group (2 modes)

- **Associate** — collection-valued `POST /<set>(<id>)/<nav>/$ref` OR
  single-valued `PATCH @odata.bind`. Cardinality detected from metadata; verb
  auto-flips with the nav picker. Polymorphic lookups (Customer, Owner,
  `regardingobjectid`) emit the correct target-disambiguated
  `<navName>@odata.bind` binding shape.
- **Disassociate** — collection-valued `DELETE /<set>(<id>)/<nav>(<targetId>)/$ref`
  OR single-valued `PATCH @odata.bind: null`. Same polymorphism story.

#### ⚡ Execute group (3 modes)

- **Execute Action** — OOB / Custom API / Custom Action invocation via
  bound `POST /<set>(<id>)/Microsoft.Dynamics.CRM.<Action>` OR unbound
  `POST /<Action>`. CSDL-driven parameter form with type-aware controls:
  EntityReference (record picker), EntityCollection (multi-record picker),
  Edm primitives (typed inputs), Enum / Status-code OptionSet (live-resolved
  picker), ComplexType (nested form). Required-first UX with optional params
  collapsed into an Accordion; Form ↔ JSON toggle via Monaco; Request body
  preview card.
- **Execute Function** — OData function calls with `?@p1=...` parameter
  aliasing OR inline literal encoding. Colorized URL preview. Same
  type-aware controls as actions.
- **Execute Workflow** — runs on-demand workflows via
  `POST /workflows(<id>)/Microsoft.Dynamics.CRM.ExecuteWorkflow` with the
  bound record id. Live picker over activated, on-demand process workflows
  (`statecode eq 1 and type eq 1 and ondemand eq true and category eq 0`).

#### 💾 Binary group (3 modes)

- **Manage File Data** — File-column upload / download / delete pipelines:
  - 3 upload paths — single-request `PATCH /<set>(<id>)/<col>` with
    `x-ms-file-name` header, chunked PATCH with session token in `Location`
    response, or `InitializeFileBlocksUpload` / `UploadBlock` /
    `CommitFileBlocksUpload` messages.
  - 4 download paths — single GET `/$value`, ranged GET with `Range` header,
    message-based block download via `InitializeFileBlocksDownload` /
    `DownloadBlock`, or `GetFileSasUrl` function for SAS-URL handoff.
  - Delete via `DELETE /<set>(<id>)/<col>` OR `POST /DeleteFile {FileId}`.
- **Manage Image Data** — Image columns with thumbnail vs full-size download
  (`?size=full`), `CanStoreFullImage` awareness, and explicit pointers to
  the file-column messages path for full-size uploads. Three doc-equivalent
  delete forms (DELETE, PATCH-to-null, PUT-with-null).
- **Manage Attachment / Note** — `activitymimeattachment.body` and
  `annotation.documentbody`. Inline base 64 (≲ 4 MB) OR
  `Initialize{Attachment|Annotation}BlocksUpload` /
  `Commit{Attachment|Annotation}BlocksUpload` messages. Single-request
  GET `/$value` OR message-based block download.

#### 🔍 Cross-cutting capabilities

- **Live Dataverse metadata** — tables, columns, relationships, alternate
  keys, action / function CSDL all fetched lazily from the host's
  `dataverseAPI`. Singleton TTL cache (5-minute TTL) with in-flight promise
  dedup. Type-specific column properties (Targets, OptionSet, MaxLength,
  MinValue / MaxValue, Format) loaded on-demand per column.
- **Antipattern advisories** — real-time guidance per the Microsoft
  query-antipatterns doc: `FilteringOnCalculatedColumns`,
  `LargeAmountOfLogicalAttributes`, large-text `contains`, leading-wildcard
  rewrites. Surfaces inline at the source AND aggregates into an Advisory
  drawer in the URL bar.
- **Polymorphic lookup support** — Customer, Owner, and multi-target
  lookups (`regardingobjectid`) emit the correct target-disambiguated
  `<navName>@odata.bind` binding shape automatically.
- **Code generators** — every mode emits the same request as
  `Xrm.WebApi.online.execute` (single + `executeMultiple` batch), `fetch`,
  `curl`, and PowerShell `Invoke-RestMethod` snippets. Multi-request
  pipelines (chunked uploads, message-based binary ops) render as full
  ordered sequences in the Code tab.
- **Saved-request library** — per-org localStorage persistence (50-entry
  cap per org), auto-suggested names, rename / delete / overwrite, dirty-state
  detection. Cross-tab sync via `storage` event listener.
- **Results grid** — virtualized via `@fluentui-contrib/react-data-grid-react-window`
  for smooth scrolling of large result sets. Per-AttributeTypeCode cell
  renderers (Switch for Boolean, Badge for Picklist / Status, lookup
  formatted-value rendering, locale-formatted dates / numbers / money).
  Sort indicators driven by `$orderby` state. Search bar + density toggle.
  Infinite scroll fires `onLoadMore` 10 rows from the bottom; Retrieve-All
  walks every page until exhausted.
- **PPTB host integration** — `isEmbedded()`-aware UI: theme syncs with the
  host, environment-aware URLs resolve via `getSession().environmentUrl`,
  header / theme controls hide when the host owns them. Bypass headers
  surface in the Code tab for external execution.
- **Standalone mode** — runs in a regular browser via `npm run dev`. Execute
  returns `501 Not Implemented` (no host bridge), but the URL bar, body
  preview, and Code tab work end-to-end.

#### 🎨 UX

- **Required-first parameter UX** — required action params / required
  create columns surface in the main pane; optional params collapse into
  an Accordion with a "N of M set" badge.
- **Two-dropdown cascading picker** — Scope (group / bound entity) +
  Operation for Execute modes. Picking a scope filters the operation list;
  scope change cascades-resets the operation.
- **Form ↔ JSON toggle** — Monaco-backed JSON editor for action params and
  create / update bodies. Two-way live sync via `setValues` so edits in JSON
  reflect in the form and vice versa.
- **Diff view in Update mode** — fetches the current record on entry,
  shows changed rows (before → after) with color coding, unchanged rows
  collapsed under a "Show unchanged" toggle.
- **Operation-specific result narratives** — `WriteResultCard` renders
  per-operation success / failure copy (Create returns 204 with
  `OData-EntityId`; Delete returns 204; binary ops return 501 framed as
  "request built, run externally"; ETag-stale 412 has its own narrative).
- **Sidebar with clause tree** — every mode has the same shell: target
  pane on the left, active editor in the center, URL bar + Execute on the
  bottom. Sections expand into clause sub-items with dirty markers.
- **Theme support** — light + dark via Fluent v9 themes, with studio-owned
  CSS variables for brand / method palettes. Group-color cascade overrides
  Fluent's brand tokens per mode group (Read = blue, Write = green,
  Relate = purple, Execute = orange, Binary = brown-red).

### 🏗️ Technical

- **TypeScript strict mode** — zero `any` (with eslint-escape-hatched
  exceptions documented inline). Discriminated unions for `ColumnMeta`,
  `FileUploadMethod`, `Advisory`, every clause node type.
- **Vite 7** — production build splits Monaco into its own chunk
  (~3.7 MB on disk, lazy-loaded on Code-tab open). Main bundle ~1.87 MB
  raw / 494 KB gzipped.
- **Fluent UI v9 (9.73)** — no v8 / v0 fallbacks. Custom `triggerButton`
  styling for `NavPathColumnPicker` uses `!important` over Fluent tokens
  to defeat the appearance-variant specificity in dark mode.
- **`@pptb/types` 1.2.1** — official Power Platform ToolBox type
  definitions; `features.minAPI: "1.2.0"` in the manifest pins the
  required host version.
- **AGPL-3.0-only** license.

### 🔗 References

- [Use the Dataverse Web API](https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/overview)
- [Query antipatterns](https://learn.microsoft.com/en-us/power-apps/developer/data-platform/query-antipatterns)
- [File column data](https://learn.microsoft.com/en-us/power-apps/developer/data-platform/file-column-data)
- [Image column data](https://learn.microsoft.com/en-us/power-apps/developer/data-platform/image-column-data)
- [Attachment & Note file data](https://learn.microsoft.com/en-us/power-apps/developer/data-platform/attachment-annotation-files)
- [Use Web API actions](https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/use-web-api-actions)
- [Use Web API functions](https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/use-web-api-functions)
- [Associate / Disassociate](https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/associate-disassociate-entities-using-web-api)
- [Retrieve and execute predefined queries](https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/retrieve-and-execute-predefined-queries)
