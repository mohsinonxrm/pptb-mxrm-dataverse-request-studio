# Changelog

All notable changes to Dataverse Request Studio will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### ✨ Added

- **Writes ask before they're sent** ([#241]). The confirmation names the
  target: the environment (connection name, URL and type, with Production
  highlighted), the operation, the table and the scope. Execute and Ctrl+Enter
  both go through it.
  - Merge, Execute Action / Custom API / Custom Action, Execute Workflow, Bulk
    Write, `$batch` and ExecuteMultiple with writes, and multi-target
    Associate / Disassociate always ask.
  - Create, Update, Upsert, Delete and single-target Associate / Disassociate
    ask on a Production connection.
  - On Production, and for more than 100 records (or batch writes) anywhere,
    you type the table (or operation) name before the button enables.
  - If the connection changes while the confirmation is open, nothing is sent.
  - The Bulk Delete Job confirmation names the environment too.
- **The header shows the connection and its environment type** inside
  ToolBox too, with Production in red ([#241]).
- **Production safe mode** in Settings → Safety, on by default ([#241]).
- **README: "What this tool changes"** lists every request that changes data,
  when it asks, and what DRS stores locally ([#241]).
- **Read, Execute Function, Associate / Disassociate and Bulk Delete Job
  requests send their headers** ([#109]). The Headers pane says which rows
  Execute sends and marks the rest as Code tab only. Later pages of a
  Retrieve Multiple send the first page's headers, so they run as the same
  user. It's an allow-list, so a plain read still sends none and never asks
  for consent:
  - impersonation (`MSCRMCallerID`, and a new `CallerObjectId` preset for
    the Microsoft Entra object id) in all of them;
  - `If-None-Match` and `Consistency` on reads. A matching etag shows as
    **304 Not Modified**, not as an error;
  - the bypass headers on Associate / Disassociate.
  The page size's `Prefer` still keeps ToolBox's
  `odata.include-annotations="*"`; other preferences stay in the Code tab.
- **Impersonation is never dropped silently** ([#109]).
  - On a ToolBox older than 1.2.7, which can't send headers, a request with
    an impersonation header isn't sent, instead of running as you. The same
    goes for Associate / Disassociate with a bypass header.
  - Both impersonation headers at once, or the presets' all-zero
    placeholder, aren't sent either.
  - A declined header prompt on Associate / Disassociate asks once and sends
    nothing, rather than falling back to a request without your headers.
- **Execute Custom API and Custom Action list your own operations first**
  ([#248]). A Dynamics org has well over a thousand of them, nearly all
  Microsoft's. The Operation list now groups them:
  - **Your org & ISVs** first, **Microsoft** below. Nothing is hidden, and
    search covers both groups (publisher names too).
  - Each one shows its solution's publisher, and so does the summary card.
  - The Scope list counts your own operations per scope ("12 ops, 2 yours").
  - Execute Action (WinOpportunity, QualifyLead, …) is unchanged.

### 🐛 Fixed

- **The sidebar works with a screen reader and the keyboard** ([#243]). It
  announced itself as a tree, but arrow keys did nothing, nested clauses
  were read as top-level ones, and every parent said "expanded" although
  nothing collapses. It's now a "Request structure" navigation landmark:
  - each section has a heading and is a list named by it, and nested
    clauses (an `$expand` relationship's clauses, required parameters, Bulk
    Write rows) are nested lists, so their level is read;
  - Tab moves between clauses, Enter or Space opens one, and the open one
    is marked as current;
  - the chevrons, which looked collapsible but weren't, are gone;
  - the "modified" marker and the ✓ badge are read out;
  - keyboard focus shows a ring inside the sidebar;
  - the table name after an expanded relationship (`primarycontactid →
    Contact`) has enough contrast to read.
- **Small labels are readable again** ([#330]). Badges such as "req",
  "null", "File", "max 32 MB", "soon", "lookup" and the Results count used
  Fluent's extra-small size, a 10px dot, so their text spilled out or was
  cut off. They now use the small size, as the sidebar badges do since
  [#200].
- **Expired metadata is refetched, not served as fresh** ([#334]). The table
  list, and the columns of a table not yet loaded this session, are cached
  for an hour. When that hour is up and they're fetched again, other pickers
  now wait for the fetch instead of getting the old data, and a failed fetch
  is retried on the next use instead of the old data counting as fresh for
  another hour.
- **Choosing a file no longer reads it into memory** ([#222]). Manage File,
  Image and Attachment Data keep the file's name, size and type only: Execute
  can't send binary content from inside ToolBox, and the Code tab uses a
  placeholder for the bytes, so a large file used to freeze the tool for
  nothing. Manage Image previews the chosen image through an object URL (up
  to 5 MB). A pasted base64 body is capped at 10 MB, with a message.
- **Non-ASCII file names work in generated upload code** ([#222]). The
  `x-ms-file-name` header only takes ASCII (and `fetch` / XHR throw on such
  a header), so a name like `报告.pdf` or `Résumé.docx` travels as the
  `x-ms-file-name` query parameter instead, as Microsoft documents, in both
  the single-request and chunked uploads.
- **Invalidating one table's metadata now refetches it** ([#335]). The next
  read used to return the old table from the in-memory registry. Nothing
  invalidates a single table yet, so this changes nothing you can see today.
- **"Advanced Find only" works, and is off by default** ([#44]), matching
  FetchXML Studio:
  - Off: the table, column and relationship pickers offer everything.
  - On: they offer only what Advanced Find shows (`IsValidForAdvancedFind`).
    Image and File columns are always offered.
  - Metadata now loads unfiltered, and only the pickers are scoped, so a
    column already in a request keeps working either way. Before, the switch
    did nothing, and columns and relationships that Advanced Find hides were
    never offered.
  - The read pickers still never offer columns that can't be read, and new
    sorts, filter rules and functions start on a column the picker offers.
  - The Create form doesn't treat columns Advanced Find hides (such as
    `owneridtype`) as required, as before.
  - A setting saved by an earlier version is reset to off once.
  - Loading a saved request, pasting a URL and a batch item's entity set
    check the table against every table, so a saved request for a table
    outside the Query Scope loads instead of being reported missing.
- **Ids and names are checked before they go into a URL** ([#221]). A record
  id must be a GUID, and a navigation property or column name a plain name,
  before Execute enables. This applies in Associate, Disassociate, Execute
  Action / Function / Workflow, Predefined Query and the file, image and
  attachment modes. The other modes already checked. The executor refuses
  them too, so a value like `x)/other_nav` can't send a different request.
- **A query function on a related table's column is caught inside `any` /
  `all` too** ([#336]). Dataverse rejects a function such as `EqualUserId`
  whose column runs through a relationship
  (`PropertyName='parentcustomerid_account/ownerid'`). Outside a lambda
  this already blocked Execute with an explanation; inside one it wasn't
  checked, so the request was sent and failed.
- **Filter values are typed** ([#221]). A number, GUID or date column's
  value goes into `$filter` unquoted only when it is one. Anything else is
  quoted and escaped, and Execute is blocked with a message naming the value,
  instead of a Dataverse 400. The DateTime fallback now escapes quotes. The
  same applies to the Dataverse query functions (`LastXDays`, `Under`, …),
  `$expand` inner filters and Bulk Delete Job. Filter column, lambda and
  `$expand` names must be names too. EntityName columns (`objecttypecode`)
  are now compared as strings, as the Web API expects.
- **Execute Function parameters are typed the same way** ([#221]). GUID and
  date values go into the URL unquoted only when they are one, and record
  references are escaped.
- **Imported saved requests are checked deeply** ([#221]). Entries whose
  ids, names or query trees would change the URL are skipped, and the import
  status says which and why. Copying or moving from another connection's
  library applies the same check.
- **Retrieve Single no longer turns on `If-None-Match: null`** by default
  ([#109]). Execute sends that header now, and `null` never matches an etag;
  saved requests that have it still run without it.
- **Time zone independent times in a DST gap keep their digits** ([#337]).
  Such a time is stored as typed, with no conversion. When it fell in the
  hour your clocks skip (02:30 on a spring-forward day), the results grid and
  the Update / Merge comparison showed it an hour later, and changing only
  its date, or just clicking into the time picker and out, sent the later
  time. The time picker itself still shows it an hour on, because it can
  only hold a local time.
- **Exported URLs encode spaces** ([#221]). Postman, Insomnia and Bruno
  collections, cURL / fetch / XHR / C# / PowerShell code and the HTTP-action
  URIs write `%20`. The URL bar still shows spaces.

### 🔧 Internal

- **Tests for the write path** ([#235]): 272 new tests, 804 in all.
  - A parity test checks, mode by mode, that Execute sends exactly what the
    Code tab shows: method, URL, body and headers, apart from the documented
    differences (headers the host builds, the reads' allow-list and page
    size).
  - The write executors' refusals: PUT of one column, alternate-key Upsert,
    single-property Delete, unsafe ids, and impersonation on an older
    ToolBox.
  - Modules that had no tests: date and time values, the bypass headers,
    workflow arguments, request validation, the query advisories, mode
    drafts, the metadata cache and provider, and the ToolBox bridge's events.

[#330]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/330
[#334]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/334
[#235]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/235
[#248]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/248
[#335]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/335
[#222]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/222
[#44]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/44
[#336]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/336
[#109]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/109
[#221]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/221
[#337]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/337
[#241]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/241

## [1.4.0-beta.3] - 2026-10-27

The third 1.4.0 beta fixes what a full regression pass on beta.2 turned up
([#304]). The headline fix: Manage Image Data and Manage File Data can pick
columns again. The rest are smaller fixes to error handling, wording and
the saved-request library.

Still a **prerelease** on the npm `beta` tag, with the gold Insider icon.
Every fix was verified in Power Platform ToolBox 1.2.8 against a live
Dataverse environment.

### 🐛 Fixed

- **Manage Image Data and Manage File Data list their columns again**
  ([#305]). Image and File columns report the type `Virtual` and aren't
  Advanced Find-eligible, so DRS dropped them, and both modes said the table
  had no Image or File columns, even Account's `entityimage`.
  - They're now loaded in the same metadata request, and `entityimage` can
    also be used in `$select`.
  - Each column's size limit and full-size image setting come from the
    environment, so upload size checks use the column's real limit, and the
    table's primary image is marked as thumbnail-only.
  - **Order by** doesn't offer Image or File columns, which can't be sorted
    on.
  - A required Image or File column no longer blocks Create or Bulk Write.
    Those columns are uploaded separately and aren't in the field set.
- **Execute Workflow can retry a failed workflow list** ([#306]). If loading
  the on-demand workflows fails (for example a brief `HTTP 401` from the
  host), the error has a **Retry** button, and the picker says the workflows
  didn't load instead of "No activated workflows in this environment".
- **Associate and Disassociate name the right HTTP verb** ([#307]).
  Disassociate's targets pane said each target fires a POST; it fires a
  DELETE. Single-valued navigation properties said a PUT sets the lookup;
  it's a PATCH with `<nav>@odata.bind`. Disassociate also showed the "Target
  records" heading twice.
- **Saved-request library** ([#308]):
  - With **All modes** off, an empty list says how many requests are saved
    in other modes, instead of "No saved requests yet".
  - Loading a saved request over a saved request you've edited asks before
    discarding your changes.
  - After you delete the saved entry a request came from, **Save** works
    again without editing the request first.
- **Smaller wording fixes** ([#309]):
  - Execute Action and Execute Function results point at the **JSON** tab.
  - The response properties of an action or Custom API come from the
    environment's `$metadata`, including inherited properties, instead of a
    note about a "mock CSDL".
  - Recent runs show a row count only for reads ("1 row", "6 rows"), not
    "204 · 1 rows" for writes. Associate and Disassociate say how many
    links were done ("3 of 5 linked"), and Bulk Write how many rows were
    written.
  - The **Learn more** link on an unknown-column error goes to *Select
    columns*, or *Order rows* for `$orderby`.

[#304]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/304
[#305]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/305
[#306]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/306
[#307]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/307
[#308]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/308
[#309]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/309

## [1.4.0-beta.2] - 2026-10-24

The second 1.4.0 beta fixes the issues found while testing beta.1 and
finishes the v1.4.0 milestone:
- paged reads now honour your page size;
- Retrieve all stays fast on very large tables;
- Execute Action and Execute Workflow send their headers (impersonation,
  bypass);
- malformed ETags are caught before they're sent;
- a startup race that could leave DRS without a connection is fixed.

Still a **prerelease** on the npm `beta` tag, with the gold Insider icon.
Every fix was verified in Power Platform ToolBox 1.2.8 against a live
Dataverse environment.

### 🛡️ Security

- **DOMPurify 3.4.17 in the Monaco editor.** Monaco bundles its own copy of
  DOMPurify, and 1.4.0-beta.1 shipped 3.1.7, which open Dependabot alerts
  (up to high severity) cover. The build now swaps in the npm package at
  3.4.17, and CI and the release workflow check the bundled version.
  Command Bar Studio and FetchXML Studio use the same fix.

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
- **Retrieve all stays responsive on large tables** ([#223]).
  - Pages reach the grid in batches. The size is added up page by page,
    instead of re-measuring everything loaded after every page, which slowed
    down with every page.
  - It pauses every 50,000 rows (or 1,000 pages) and says how many rows are
    loaded. **Continue** loads the next batch. It used to stop silently after
    200 pages, so the grid and exports could look complete when they weren't.
  - Response sizes are counted in UTF-8 bytes, so text in other scripts is
    no longer undercounted.
- **Recent runs keeps up with paging** ([#211]): after scrolling or Retrieve
  all, the entry shows the rows loaded, with "more available" or "all pages".
- **Execute Action and Execute Workflow send your request headers** ([#266]).
  On ToolBox 1.2.7+, the enabled rows in their Headers pane (for example
  `MSCRMCallerID` / `CallerObjectId` impersonation, or `MSCRM.*` bypass
  headers) now go out with the request, after ToolBox asks you to approve
  them. They used to be dropped, so an action or workflow ran as you without
  saying so. Reads (apart from the page size, [#219]), relate, Execute
  Function, Bulk Delete and the file / image / attachment modes still run
  with ToolBox's own headers ([#109]).
- **A malformed ETag is caught before it's sent** ([#264]). Update,
  Upsert and Delete send the **Specific ETag** trimmed. They block Execute,
  with the reason, unless it is a quoted tag such as `W/"12345678"` (or
  `*`). Before, a bare version number, curly quotes, JSON escaping
  (`W/\"…\"`) or stray text went out and Dataverse answered with a 412 or
  400 that didn't point back at the value. **Fix** corrects those paste
  mistakes in one click. `$batch` parts and exported collections send the
  trimmed value too.
- **Error and result titles name the right source** ([#209]).
  - When ToolBox reports no HTTP status, the error card says the message
    comes from Dataverse only when it carries a Dataverse `0x8…` code. It
    names ToolBox only for ToolBox's own messages, such as its
    response-size limit, and claims neither source otherwise. The size hint
    also mentions `$top` for requests without a page size.
  - Create results name the record by the primary name the response
    returns, so a new contact shows its full name instead of
    "record <id>".
- **"N populated" counts only values that will be sent** ([#210]). The badge
  and the red "missing" marker on required rows ignore the empty values a
  table pre-fills for its required columns, the same way the required count
  does. Before, a new contact could read "2 populated" with Last Name and
  Owner still empty. The **Set** filter still lists those pre-filled rows, so
  they stay visible to fill in.
- **No more "No active Dataverse connection" after the connection is
  picked** ([#281]). When ToolBox bound the connection while DRS was still
  starting up, DRS could miss it and stay on that error until it was
  reopened. It now listens for the connection from the start, and Retry asks
  ToolBox again.
- **Code → Editor no longer fakes a run** ([#240]). Its **Run (mock)**
  button ignored the edited code and always showed the same two sample
  account rows, which read like "your code works". The button is gone: the
  editor is for editing and copying, and **Execute** sends the request the
  builder shows.
- **The "…" menu has no dead items, and the counts are right** ([#265]).
  **Duplicate request** and **Export OpenAPI** did nothing and are removed.
  OpenAPI is one of the Code tab's formats, with **Save**. The mode picker and
  the Code tab now count their modes (23) and formats (10) instead of saying
  18 and 8.
- **Docs caught up** ([#237]). CONTRIBUTING lists `npm test` and the CI
  checks, and describes the tag-driven release (RELEASING.md) instead of a
  manual publish. README describes the per-org metadata caches and the
  Retrieve all pause.

[#109]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/109
[#209]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/209
[#210]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/210
[#211]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/211
[#219]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/219
[#223]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/223
[#237]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/237
[#240]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/240
[#264]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/264
[#265]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/265
[#266]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/266
[#281]: https://github.com/mohsinonxrm/pptb-mxrm-dataverse-request-studio/issues/281

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
