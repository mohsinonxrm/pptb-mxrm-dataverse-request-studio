# Changelog

All notable changes to Dataverse Request Studio will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.1.3] - 2026-06-10

### ✨ Added

- **Auto-switch to Builder tab on sidebar item click**
  ([#32](https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/32))
  — clicking any sidebar nav item while on the Code or Results tab now
  automatically switches the active tab back to Builder, making the selected
  editor panel immediately visible. Previously the tab had to be switched
  manually first. Applies across all 17 modes.

### 🔧 Internal

- **Prettier added as dev tooling** — `.prettierrc` (single quotes, 2-space
  indent, semis, `printWidth: 100`), `.prettierignore`, `.gitattributes` (LF
  line endings repo-wide), and `.vscode/settings.json` (Prettier as default
  formatter with `formatOnSave`). Added `format` and `format:check` npm
  scripts. Full codebase normalised in one pass to eliminate quote-style
  inconsistency introduced by a recent VS Code built-in formatter change.

## [1.1.2] - 2026-06-08

### 🐛 Fixed

- **`refreshAll()` race condition** (`metadataProvider.ts`) — the Settings
  "Refresh metadata" button now drains any in-flight `buildTable` promises via
  `Promise.allSettled` before clearing the cache. Previously an orphaned fetch
  that settled after the clear could call `__registerLiveTable` with
  pre-refresh data, silently overwriting the fresh registration.

- **Silent bad-hop in OData parse validation** (`odataParser.ts`) — when the
  parser walks a nav-path `$filter` column and encounters an unknown or
  non-ManyToOne navigation segment, it now emits an actionable warning (naming
  the bad segment and its owner table) instead of silently skipping validation.
  A separate warning is emitted when the related entity's metadata times out
  during the walk.

- **`resolveNavPath` leaf lookup now matches `oDataName`** (`mock/metadata.ts`)
  — the leaf-column search now also checks `ColumnMeta.oDataName` (e.g.
  `_primarycontactid_value`) in addition to `logicalName`. This closes a gap
  where a round-tripped or pasted OData `$filter` that preserved the lookup
  wire form as the path leaf would resolve to `undefined`, causing the `$filter`
  encoder to fall back to string-quoting a value that should be unquoted
  (GUID / integer).

## [1.1.1] - 2026-06-08

### 🐛 Fixed

- **Custom OptionSet (and other typed) values incorrectly quoted as strings in
  `$filter` over a lookup navigation** ([#33](https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/33)).
  Filtering on a related-entity column via a nav path — e.g.
  `msdyn_opportunityid/abc_salesstage eq 809770000` — emitted the value as a
  quoted `Edm.String` (`'809770000'`), which Dataverse rejects with HTTP 500
  / `0x80060888`. Standard columns appeared to work only by coincidence (the
  root entity happened to own a same-named column). Root cause: the `$filter`
  encoder resolved the leaf column against the **root** table instead of
  walking the relationship to the **related** entity, so the leaf's
  `AttributeType` never drove the quoting decision. The encoder now resolves
  nav-path leaves against the correct related entity, so numeric/boolean/choice
  types are emitted unquoted and strings/GUIDs are quoted, consistently.

### ✨ Added / Changed (architecture)

- **Single, metadata-driven nav-path resolver (`resolveNavPath`)** in
  `mock/metadata.ts` — now the one source of truth for resolving a
  slash-delimited path (`nav/.../leaf`, alias-aware) to its leaf `ColumnMeta`,
  owner table, visited chain, and first not-yet-loaded hop. The `$filter`
  encoder (`colLookup`), the filter editor's leaf/pending-target resolution,
  and the `$orderby` warnings all route through it, eliminating duplicated
  walk logic that had drifted out of sync.
- **Related-entity metadata pre-warming (`useWarmReferencedTables`)** — the
  Retrieve Multiple / Retrieve Single modes now walk the `$filter`, `$expand`,
  `$apply`, and `$orderby` trees and pre-fetch the related entities they
  reference. This makes correct type resolution independent of whether the
  user opened the relevant editor — covering **saved-request reload**,
  **pasted OData URLs**, and **direct Execute**. Targeted (only referenced
  entities) and self-healing across multi-hop paths, so it stays clear of the
  100-concurrent-request cap.
- **Paste-OData parsing** now validates and pre-loads nav-path filter leaves
  instead of skipping them.
- **Metadata cache TTL raised from 5 minutes to 1 hour**, with a new
  **Settings → Refresh metadata** button (`metadata.refreshAll`) to force a
  fresh pull after publishing schema changes mid-session.

## [1.1.0] - 2026-05-25

### ✨ Added

- **Scope-setup guidance banner** — when the tool loads with the entity scope
  unconfigured (i.e. `entityScopeMode` is `publisher-solution` or
  `solution-only` but no solutions have been selected), a contextual
  `MessageBar` (warning intent) is shown at the top of the Target pane across
  **all 14 modes** that use `TargetEditor`. The banner provides a mode-specific
  hint and an **Open Settings** button that opens the Settings drawer directly.
  It disappears automatically once the user selects a publisher / solution (or
  switches to *All Entities*) — no dismiss cookie required, because the
  condition is purely reactive to the persisted `DisplaySettings` state.

- **`OpenSettingsProvider` / `useOpenSettings` / `useRegisterOpenSettings`**
  (`src/host/useOpenSettings.tsx`) — new React context that provides a
  stable `() => void` callback to open the Settings drawer from any component
  in the tree without prop-drilling. Uses a ref-based registration pattern
  (zero provider re-renders). `FrameHeader` registers the callback on mount;
  any consumer calls `useOpenSettings()` to get it.

- **`useScopedEntities` now exposes `needsSetup: boolean` and
  `scopeMode: EntityScopeMode`** — consumers can detect the
  unconfigured-scope condition and tailor guidance text accordingly.

### 🔍 Settings persistence confirmation

- Verified that all `DisplaySettings` (scope mode, selected publisher /
  solution IDs, display flags) are persisted via `window.toolboxAPI.settings`
  under the namespaced key `pptb-dataverse-request-studio:displaySettings`.
  The prefix is distinct from other PPTB tools (e.g. FetchXML Studio). No
  changes to the persistence layer were required.

---

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
