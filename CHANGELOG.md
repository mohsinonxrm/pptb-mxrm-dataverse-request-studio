# Changelog

All notable changes to Dataverse Request Studio will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### 🐛 Fixed

- **A page size no longer turns `$top` into "up to 5,000 rows"** ([#219]).
  Retrieve Multiple used to drop `$top` whenever `Prefer: odata.maxpagesize`
  was set, but Execute never sent that Prefer.
  - On ToolBox 1.2.7+, Execute now sends the page size you set for
    Retrieve Multiple, Predefined Query or Retrieve NextLink. In Retrieve
    Multiple, scrolling and Retrieve all ask for the same size on every
    page. ToolBox asks you to approve the `Prefer` header the first time.
    Execute still asks for all annotations, so the grid renders as before.
    Retrieve NextLink no longer starts with a page size of 50, so a pasted
    link keeps the size it was made with unless you set one.
  - On older ToolBox versions, Execute keeps `$top`, and the hints suggest
    `$top` instead of a page size that can't be sent.
  - `Prefer` now puts `odata.maxpagesize` last: in the Code tab, in
    `$batch` parts and in exported collections. Dataverse rejects it before
    `odata.include-annotations` with `0x80060888` "Invalid prefer header",
    so requests with a page size and annotations used to fail.

## [1.4.0-beta.1] - 2026-10-21

The biggest release so far: four new modes (Bulk Write, Bulk Delete Job,
`$batch`, ExecuteMultiple), request headers sent through Power Platform ToolBox
1.2.7+, a reworked results grid, a portable saved-request library, and dozens of
fixes. The work planned for v1.2.0, v1.3.0 and v1.4.0 ships together, so
there are no separate 1.2.x or 1.3.x releases; the few items still open move
to later milestones.

This is a **prerelease** on the npm `beta` tag. It uses the gold Insider icon.
Every change was verified end-to-end in Power Platform ToolBox against live
Dataverse environments.

### ✨ Added

- **Bulk Write mode** ([#8]): `CreateMultiple` / `UpdateMultiple` /
  `UpsertMultiple` / `DeleteMultiple` over a grid of rows, with a per-table
  support check.
- **Bulk Delete Job mode** ([#10]): a `BulkDelete` system job built from a
  query, started now or scheduled, with live job status.
- **`$batch` Request mode** ([#9]): change sets (atomic, with `$n` Content-ID
  references) and independent requests, continue-on-error, "Add last request"
  from the read, write, relate and execute modes, and per-part results. Recent runs show the batch status and
  a per-part summary ([#204]).
- **ExecuteMultiple mode** ([#11]): independent requests with
  continue-on-error, on the `$batch` engine, with SDK
  `ExecuteMultipleRequest` (C#), raw `$batch` and PPTB `dataverseAPI` code.
- **Request headers through Power Platform ToolBox 1.2.7+** ([#43], [#23]):
  Create, Update, Upsert, Delete, Merge and Bulk Write now send their preconditions (`If-Match` / `If-None-Match`), `MSCRM.*`
  bypass headers, duplicate detection and impersonation are now sent. ToolBox
  asks you to approve additional headers once; choose **Approve for this
  tool** so it doesn't ask again. On older hosts, requests that depend on a
  header stay blocked with an explanation, as in 1.1.4.
- **Results grid** ([#12], [#15], [#16], [#17]): CSV (as shown or raw) and
  JSON export, persisted column widths, extra-small / small / medium density,
  a per-column filter row, and an indexed row search.
- **Saved-request library**:
  - versioned, with Export / Import to a JSON file ([#105]);
  - export as a Postman, Insomnia or Bruno collection ([#27]);
  - copy or move saved requests from other connections ([#6]);
  - an "All modes" view, and clearer name-clash messages ([#125]).
- **OpenAPI 3 document** for the current request in the Code tab ([#28]).
- **Click a recent run to restore it** into its mode ([#2]).
- **Retrieve this record** from the Create, Update, Upsert and Merge result
  card ([#3]).
- **Execute Workflow shows the workflow's declared input arguments** ([#4]).
- **Predefined Query over a row's related records** ([#5]).
- **Address shortcut chips** in the `$filter` builder ([#7]).
- **EntitySpecific parameters** can reference an existing record or build a
  new instance ([#1]).
- **Response-size hint**, and outer `$top` with a nested collection `$expand`
  no longer blocks ([#126]).

### 🔄 Changed

- **Execute Custom API / Custom Action list every operation by type**,
  Microsoft's own included, classified from Dataverse data (`customapi`,
  `workflow` + `sdkmessage`); the name is only a fallback when those lookups
  fail ([#54]).
- **Update sends `If-Match: *` by default** on ToolBox 1.2.7+, so an Update of
  a record that doesn't exist returns 404 instead of creating it ([#43]). The
  first Update shows ToolBox's header approval prompt.
- **Each mode keeps its unsaved request** when you switch modes ([#101]).
- **Switching the ToolBox connection resets the prepared request** and
  reloads metadata, caches and lists for the new environment ([#52], [#93],
  [#107]).
- **In-app dialogs** replace the browser's native confirm / alert boxes
  ([#111]).
- **Errors show the real status** and no made-up response headers ([#59]);
  Electron's IPC wrapper is stripped from host messages ([#188]).
- **Multi-target Associate / Disassociate report each target**, including
  "Partial success" ([#99]).
- **Faster startup:** Monaco and the less-used modes load on demand ([#22]);
  long picker lists are capped and filtered lazily ([#21]).
- **Render errors are contained** to the mode or pane that failed ([#19]).

### 🛡️ Security and data safety

- **Credential headers are never sent, saved or exported** ([#217]):
  `Authorization`, `Proxy-Authorization` and `Cookie` are stripped from saved
  requests (including ones saved by earlier versions), library exports, the
  Code tab, OpenAPI and `$batch` parts. Impersonation ids become variables in
  Postman / Insomnia / Bruno exports.
- **CSV export and Copy as TSV neutralise spreadsheet formulas** ([#239]).
- **Update can't create a record on ToolBox < 1.2.7** ([#218]): DRS checks the
  record exists first.
- **Single-valued Associate / Disassociate can't create the source record**
  ([#254]).
- **The saved-request library is never overwritten when it can't be read**
  ([#220]); damaged or newer-version items are kept as they are. Import and the
  other-connections view no longer drop `$batch`, ExecuteMultiple and Bulk
  requests.
- **Bulk Write sends its bypass and other headers** on ToolBox 1.2.7+, so
  "bypass custom business logic" really skips plug-ins ([#257]). Bulk Write
  and Bulk Delete Job errors show the real status.
- **The library never evicts a request silently** at its 200-request cap
  ([#100]).

### 🐛 Fixed

**Queries**

- `$orderby` on a lookup uses `_<col>_value` ([#62]).
- `$expand` always emits an inner `$filter` and nested `$expand` ([#63]).
- `not` is kept on negated groups when parsing a `$filter` ([#117]), and a
  double `not` on `contains` / `startswith` / `endswith` cancels out ([#207]).
- Record search text is percent-encoded and wildcards match literally ([#95]).
- Stale paging no longer overwrites newer results; page errors are shown
  ([#98]).
- The footer and Recent runs count the same rows for `$expand` queries
  ([#122]).
- Date filters and grid cells show the picked date west of UTC ([#106]).

**Writes**

- DateTime fields keep the date and time you pick for each behaviour ([#92]).
- The JSON body editor keeps what you type, including nulls and unknown keys
  ([#91]).
- One required-field rule for Form and JSON mode; the required count matches
  the list ([#124], [#199]).
- A choice column without a default isn't pre-filled ([#203]).
- Update's Field set no longer uses Create wording or warns about unset
  required fields ([#80]).
- Merge lists its tables from the unscoped entity list, adds `lead`, and reads
  lookups from `_<col>_value` ([#97], [#96]).
- Disassociate reads the current single-valued lookup correctly ([#57]).
- By GUID entries are normalised, an invalid one drops the picked record, and
  the mode and text survive pane switches ([#94], [#189]).
- The previous saved request isn't overwritten after switching tables
  ([#58]).
- Record-picker names refresh after DRS writes to the record ([#81]).
- The record picker can't get stuck on "Searching…" ([#104], [#121]).
- The Execute spinner clears when a request throws ([#102]).
- Settings aren't saved before they load, and earlier changes are kept
  ([#103]).
- Metadata refreshes in place and re-enriches refreshed columns ([#53]).

**Execute**

- Collection-bound operations run on the entity set, and functions run from
  the built URL ([#55]).
- Execute Action sub-modes reset on switch; a loaded snapshot's binding is
  checked ([#56]).
- Valid OData literals and escaping in Function URLs ([#61], [#110]).
- Alternate-key Upsert URL literals and encoding ([#64]).
- Parameter collections clear after Enter; an empty number isn't added as 0;
  an untyped reference starts empty ([#246]).
- The `$metadata` size-limit message no longer names a ToolBox version
  ([#113]).

**Code, copy and accessibility**

- Code tab generator correctness across formats ([#60]).
- C# snippets escape backslashes, quotes and line breaks in header values
  ([#244]).
- Copy fixes: bypass wording, 201 / 204 badges, single name in Disassociate
  ([#123]).
- Prefer and Precondition help say what ToolBox actually sends ([#108]).
- The Field set's Send-as-null switch stays clear of the value input at narrow
  widths ([#112]).
- Accessibility: labelled inputs, mode-picker state, readable sidebar badges,
  live-region reset ([#20], [#120], [#200]).
- DRS asks you to reopen it when a lazily loaded part is missing after an
  update ([#187]).
- Relative asset URLs, so the editor loads under the ToolBox VS Code extension
  ([#212]).

### 🧰 Developer

- Vitest suite for the request engine, code generators, storage and exporters
  ([#18]); CI on pull requests and pushes to `dev` / `main`, with Prettier
  and a check for root-absolute asset URLs ([#65], [#212]).
- `@pptb/types` 1.2.7 and `features.connections` declared ([#43]).
- Opt-in bundle report: `npm run build:analyze` ([#22]).
- README, CONTRIBUTING, SECURITY and templates brought up to date ([#66]).
- Dependency security updates (Vite, esbuild, PostCSS, Babel and others).

### ⚠️ Known issues

- **`$batch` continue-on-error:** ToolBox's `$batch` response parser fails when
  a failed request is followed by more requests. DRS explains it and suggests
  putting requests that might fail last ([#198]).
- **Execute modes on very large environments:** ToolBox 1.2.7 and 1.2.8 can't
  load a `$metadata` document over 16 MB. Fixed in ToolBox
  ([PowerPlatformToolBox/desktop-app#737](https://github.com/PowerPlatformToolBox/desktop-app/pull/737));
  until that ships, DRS explains the limit ([#79]).
- **ToolBox VS Code extension:** Monaco's workers may run on the main thread,
  so JSON IntelliSense may not work there ([#255]).
- **Tool icon:** ToolBox draws SVG tool icons as a single-colour mask until
  [PowerPlatformToolBox/desktop-app#738](https://github.com/PowerPlatformToolBox/desktop-app/pull/738)
  ships ([#25]).
- **Impersonation on reads:** `MSCRMCallerID` / `CallerObjectId` are only
  sent by the write modes; read, relate and execute requests run as you
  ([#219]).
- **Sidebar keyboard navigation:** the clause tree has no arrow-key model yet
  ([#243]).
- **Binary modes** (file / image / attachment) compose the request and its
  code, but Execute returns 501 because ToolBox doesn't expose those
  operations.

[#1]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/1
[#2]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/2
[#3]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/3
[#4]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/4
[#5]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/5
[#6]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/6
[#7]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/7
[#8]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/8
[#9]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/9
[#10]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/10
[#11]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/11
[#12]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/12
[#15]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/15
[#16]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/16
[#17]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/17
[#18]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/18
[#19]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/19
[#20]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/20
[#21]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/21
[#22]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/22
[#23]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/23
[#25]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/25
[#27]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/27
[#28]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/28
[#43]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/43
[#52]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/52
[#53]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/53
[#54]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/54
[#55]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/55
[#56]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/56
[#57]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/57
[#58]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/58
[#59]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/59
[#60]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/60
[#61]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/61
[#62]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/62
[#63]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/63
[#64]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/64
[#65]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/65
[#66]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/66
[#79]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/79
[#80]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/80
[#81]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/81
[#91]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/91
[#92]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/92
[#93]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/93
[#94]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/94
[#95]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/95
[#96]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/96
[#97]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/97
[#98]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/98
[#99]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/99
[#100]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/100
[#101]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/101
[#102]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/102
[#103]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/103
[#104]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/104
[#105]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/105
[#106]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/106
[#107]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/107
[#108]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/108
[#110]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/110
[#111]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/111
[#112]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/112
[#113]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/113
[#117]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/117
[#120]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/120
[#121]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/121
[#122]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/122
[#123]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/123
[#124]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/124
[#125]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/125
[#126]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/126
[#187]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/187
[#188]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/188
[#189]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/189
[#198]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/198
[#199]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/199
[#200]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/200
[#203]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/203
[#204]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/204
[#207]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/207
[#212]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/212
[#217]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/217
[#218]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/218
[#219]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/219
[#220]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/220
[#239]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/239
[#243]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/243
[#244]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/244
[#246]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/246
[#254]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/254
[#255]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/255
[#257]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/257

## [1.1.4] - 2026-10-17

Bug-fix release. Every fix below was verified end-to-end in Power Platform
ToolBox 1.2.7 against a live Dataverse environment.

### 🐛 Fixed

- **Table list loads on first open**
  ([#41](https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/41),
  [#44](https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/44)).
  The entity list waits for the PPTB connection instead of showing
  "0 entities" until a tab switch. Load errors and "no active connection"
  show in the Target pane with a **Retry** button. Solution scoping no longer
  caches an empty list, and "All Entities" no longer drops tables that aren't
  Advanced-Find-eligible.
- **Request headers PPTB can't send block Execute instead of being silently
  dropped**
  ([#45](https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/45)).
  An ETag precondition, Upsert "create-only" / "update-only", or any bypass
  option now disables Execute with an explanation. Before, an Upsert set to
  create-only could overwrite an existing row. The default Update
  (`If-Match: *`) still runs. Sending the headers is planned with PPTB 1.2.7
  support ([#43](https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/43)).
- **Delete confirmation shows the right record after loading a saved request**
  ([#46](https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/46)).
  The name is resolved for the saved record id: PPTB 1.2.7's bare single-record
  response is now recognised. The record picker no longer keeps the previously
  picked name.
- **Execute can't fire twice while a request is running**
  ([#47](https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/47)).
  A held or repeated Ctrl+Enter no longer creates duplicate records.
- **No crash on a literal `%`**
  ([#48](https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/48))
  in Retrieve NextLink's parsed view or Execute Function's URL preview.
- **cURL and PowerShell snippets keep `$select`, `$filter`, `$top`, `$ref`**
  ([#49](https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/49)).
  URLs and headers are single-quoted, so the shell no longer expands them.
- **JavaScript snippets are valid with quotes and reserved characters**
  ([#50](https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/50))
  — fetch, Xrm.WebApi, executeMultiple and XHR.
- **Power Automate tab no longer says "Delete a row" for a column delete**
  ([#51](https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/51)).
  Requests the connector's row actions can't express map to
  "Invoke an HTTP request".
- **URL bar and Code tab fill in as soon as a table is picked**
  ([#78](https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/78)).
- **Paste OData URL accepts an `@odata.nextLink`**
  ([#83](https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/83)).
  `$skiptoken` is dropped with a warning instead of rejecting the paste.
- **Tool icon ships at `dist/icons/icon.svg`**
  ([#25](https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/25)),
  where PPTB looks for it. Prerelease builds use the Insider icon. The icon's
  `viewBox` is cropped to the hexagon, so it fills PPTB's 20 px icon slot.

### ✨ Improved

- **Solution pickers show unique name, version and managed state**
  ([#40](https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/40))
  on a second line under each name, so patches can be told apart. Search
  matches them too.

### 📦 Packaging

- **No more `npm-shrinkwrap.json` in the package.** Installing the tool no
  longer pins every transitive dependency. `finalize-package` now runs
  `validate` → `build`. The lockfile is committed as `package-lock.json`,
  which npm never publishes, so `npm ci` works in the repo
  ([#65](https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/65)).
- **README:** the version badge now reads from npm instead of a hard-coded
  version.

### ⚠️ Known issues

- **The tool icon shows in a single colour in Power Platform ToolBox 1.2.7,
  and not at all when the tool is loaded locally.** PPTB draws every SVG tool
  icon as a single-colour mask. Colour icons will render once
  [PowerPlatformToolBox/desktop-app#738](https://github.com/PowerPlatformToolBox/desktop-app/pull/738)
  ships
  ([#25](https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/25)).
- **Execute Action / Function / Custom API** can't list operations on
  environments whose `$metadata` is larger than 16 MB. This is a regression
  in Power Platform ToolBox 1.2.7; it worked in 1.2.6
  ([#79](https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/79),
  [PowerPlatformToolBox/desktop-app#736](https://github.com/PowerPlatformToolBox/desktop-app/issues/736)).
  The Execute modes now explain the limit instead of showing the raw host
  error. The same host limit fails any single response over 16 MB. Use
  `$top` / `$select` for large tables.

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
