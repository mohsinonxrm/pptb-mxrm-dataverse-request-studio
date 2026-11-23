# #225 — `$batch` as a raw-request transport (design spike)

Status: **approved by the owner on 2026-11-22**, with the decisions below. No code changes yet.

## Decisions (2026-11-22)

- **Narrow scope approved.** Only the three refused shapes go through `rawRequest`: PUT of one column, DELETE of one column's value, and alternate-key Upsert. Everything else stays on the typed calls.
- **PR 4 is dropped.** No "Send through `$batch`" toggle for GUID-addressed writes or Retrieve Single.
- **No host-version gate.** #408 made ToolBox 1.2.8 the minimum (`minAPI` 1.2.8) and removed the `HEADER_PASS_THROUGH` probe, so every supported host has `executeBatch`. So PR 1:
  - drops `hostSupportsRawRequest`;
  - doesn't keep today's 400 refusal as a fallback in PR 2.

  Sections 2.1 and 5 below are updated to match.
- The five upstream asks (§7) are for the owner to file.
Refs #225. Related: #361 (Retry-After), #233 (change tracking), #231 (compare responses), #267 (request tests).

Sources read for this note:

- **ToolBox:** `PowerPlatformToolBox/desktop-app` at tag **v1.2.8** (commit `33649f2`), read with `git show v1.2.8:<path>`. References below are `file:line` at that tag.
- **Types:** `@pptb/types@1.2.8` from npm (`npm pack`). Its `dataverseAPI.d.ts` is byte-identical to `packages/types/dataverseAPI.d.ts` at the tag.
- **DRS:** `dev` @ `d415ffa`.
- **Dataverse docs:** [Execute batch operations](https://learn.microsoft.com/power-apps/developer/data-platform/webapi/execute-batch-operations-using-web-api) and [Service protection API limits](https://learn.microsoft.com/power-apps/developer/data-platform/api-limits).

Nothing was run against a live environment for this note. §6 lists what the first implementation PR has to confirm live; PR 2 ran those checks on 2026-11-22, and §6 records what they found. Three findings changed the design: impersonation goes on the outer `$batch` (§1.7), no precondition goes with a request on one column's URL (§2.2), and the host refuses some `%` in a path (§1.2).

---

## TL;DR

- **Recommendation: build it, narrowly.** A one-operation `executeBatch` call can send the three shapes DRS refuses today:
  - PUT of a single column;
  - single-property DELETE;
  - alternate-key Upsert.

  It returns the operation's real status code and its response headers. These three shapes need no extra request headers, so they run **without a consent prompt**.
- **Keep every other request on the typed calls.** The `$batch` path has a 16 MB response cap (typed calls: 128 MB). It drops ToolBox's own impersonation, and any header on a read triggers the consent prompt. Moving more traffic onto it would make DRS worse.
- **#361 (Retry-After) is *not* unblocked by this.** A throttled `$batch` comes back as a non-multipart 429. The host turns that into a bare error message, and `Retry-After` is not in the `batchHeaders` allow-list. This needs a ToolBox change (§7).
- **#233 (change tracking) doesn't need this.** Typed `queryData` already sends a custom `Prefer` (1.2.7+) and returns the whole body, `@odata.deltaLink` included. It can go ahead independently.
- **Binary 501s:** the message-based File / Image / Attachment pipelines are JSON-only. They don't need raw transport: typed `execute()` can already send them. What blocks them is that DRS never reads the file bytes (#222). The `/$value` and chunked-PATCH routes stay impossible, because the host only sends and receives UTF-8 strings.

---

## 1. What a single-operation `$batch` can and can't express (ToolBox 1.2.8)

### 1.1 How the host sends it

`window.dataverseAPI.executeBatch(requests, connectionTarget?, additionalHeaders?)` (types `dataverseAPI.d.ts:744-745`). The call path:

1. **Preload → IPC handler** `src/main/index.ts:2135-2144`:
   - `authorizeBatch` checks header consent (§1.6).
   - `withDataverseRequest` resolves the connection slot (`index.ts:696-706`).
   - Then it calls `dataverseManager.executeBatch`.
   - Any error is re-thrown as `Dataverse executeBatch failed: <message>`. Electron then prefixes it with `Error invoking remote method 'dataverse.executeBatch': Error: `. DRS already strips that prefix (`batchExecutor.ts` `stripIpcPrefix`, #198).
2. **`executeMultipart`** `src/main/managers/dataverseManager.ts:2055-2074`:
   - validates and encodes the parts;
   - `POST /api/data/v9.2/$batch` with `Content-Type: multipart/mixed; boundary=…` and `Accept: multipart/mixed` (`:2060`);
   - parses the multipart answer and attaches `batchHeaders` (`:2072-2073`).
3. **The outer HTTP request** goes through the same `makeHttpRequest` as every typed call (`dataverseManager.ts:871-994`), with the `$batch` response cap (`:2061`).

### 1.2 What a part can carry

| Aspect | Behaviour | Host code |
|---|---|---|
| Methods | `GET`, `POST`, `PATCH`, `PUT`, `DELETE`. Anything else throws. `executeTransaction` also refuses `GET`. | `utils/dataverseBatch.ts:161-167` |
| URL | **Relative only**: a scheme, `//` or `\` throws. Leading `/` optional; `api/data/v9.2/` is added when missing, and any other `api/data/<ver>/` throws. `..`, `.`, control characters and a nested `$batch` are rejected, also after up to 3 rounds of percent-decoding. **A decoding round that throws rejects the request** ("URL contains invalid encoding"): a literal `%` in a value (`%25`) does that when what follows, once decoded, isn't a valid escape (`z%25%C3%A9` → `z%é`, live). | `dataverseBatch.ts:106-143` |
| URL query | Split with `split("?", 2)`. **JavaScript's limit truncates**: anything after a *second* literal `?` is silently dropped. | `dataverseBatch.ts:115` |
| Property paths | Not restricted. `accounts(<id>)/name` and `accounts(accountnumber='A1')` both pass validation. | `dataverseBatch.ts:106-143` |
| Headers | Up to 64 per part, 4,096 per batch. Names must be RFC 7230 tokens of 256 characters or fewer. Values must be strings with no control characters, 8,192 characters or fewer. Duplicate names must have the same value. **Nothing is reserved**: `Prefer`, `If-Match`, `Content-Type`, `CallerObjectId` and the rest go into the part verbatim. | `dataverseBatch.ts:39-75, 177-181`, `encodeHttpPart` `:207-226` |
| Default part headers | Only `Content-Type: application/json; type=entry`, and only when there's a body and no `Content-Type` was given. **No `Accept`, `OData-Version` or `Prefer` is added to a part.** | `dataverseBatch.ts:212-215` |
| Body | A `string` is sent as is; anything else is `JSON.stringify`'d. **No binary**: there's no `Uint8Array`/`Buffer` path, and the request is written as a UTF-8 string. | `dataverseBatch.ts:182-189, 211` |
| Request size | Estimated per part, then the encoded body is capped at **16 MB**. | `dataverseBatch.ts:9, 190-194, 264-266` |
| `contentId` | Optional. It defaults to the 1-based index and must match `[A-Za-z0-9._-]{1,128}`. | `dataverseBatch.ts:16, 171-175` |
| `changesetId` (new in 1.2.8) | `executeBatch` groups parts with the same id into one change set, placed at the first part's position. Parts with no id stay independent. `executeTransaction` ignores it and wraps everything in one change set. An empty or non-string id throws. | `dataverseBatch.ts:168-170, 228-262` |
| Outer headers | The optional third argument goes on the `$batch` POST. Dataverse applies outer headers **to the batch only, not to each part** (docs: "Headers applied to the `$batch` operation don't apply to each item"). **Impersonation is the exception** (live, §6): the batch runs as its caller, so an outer `CallerObjectId` / `MSCRMCallerID` applies to every part, and one inside a part is ignored. The host also always adds `Prefer: return=representation`, which merges with any custom outer `Prefer` and has no effect on the parts. | `dataverseManager.ts:886-895` |

For a single raw request, `changesetId` doesn't matter. Use `executeBatch` with one part and no `changesetId`. It's the only one of the two calls that accepts a `GET`.

### 1.3 What comes back (`BatchResult`)

Types `dataverseAPI.d.ts:22-30`. Built by `parsePart` (`dataverseBatch.ts:302-331`):

- **`status`, `statusText`**: from the part's own `HTTP/1.1 <code> <text>` line. This is the **real** status: 201 vs 204, 304, 404 and 412 come through as numbers. The typed calls only reject with a message (§1.5).
- **`headers`**: **every** header of the part's HTTP response, names lowercased, repeated names joined with `, ` (`:270-281`). No redaction is applied on this path. The only redaction in the tag is in `src/common/sentry.ts`.
- **`body`**: `JSON.parse` of the part body, falling back to the raw string. It's `undefined` when the body is empty (`:321-329`).
- **`contentId`**: from the part's MIME headers. Dataverse usually omits it outside change sets. With one part, the host then fills in the expected id by position (`:341-344`).
- **`batchHeaders`** (new in 1.2.8): an **allow-list of three outer headers** (`x-ms-dop-hint`, `x-ms-ratelimit-burst-remaining-xrm-requests` and `x-ms-ratelimit-time-remaining-xrm-requests`), copied onto every result when present (`dataverseBatch.ts:13, 30-37`, `dataverseManager.ts:2072-2073`). **`Retry-After`, `req_id` and `x-ms-service-request-id` are not in it.**

### 1.4 Size limits ("dedicated response size limits", 1.2.8)

`src/main/constants.ts:17, 24` and `dataverseBatch.ts:9-10`:

| Response | 1.2.7 | 1.2.8 |
|---|---|---|
| Typed Web API call (`makeHttpRequest` default) | 16 MB (the batch constant) | **128 MB** (`MAX_DATAVERSE_RESPONSE_BYTES`) |
| `$metadata` (`getCSDLDocument`) | 16 MB | **256 MB** |
| `$batch` request / response | 16 MB / 16 MB | 16 MB / 16 MB (unchanged) |

The 1.2.8 change *raised* the typed limits and left `$batch` where it was. A raw `GET` of a wide 5,000-row page can exceed 16 MB, and the same request through `queryData` would succeed. The limit is enforced while reading, and the error is `Dataverse response exceeds the maximum size (16 MB)` (`dataverseManager.ts:921-927`).

### 1.5 Errors, partial failures and continue-on-error

- **Outer non-2xx with a non-multipart body** (401, 429 throttling, 5xx from the front end): `makeHttpRequest` rejects with `Error("<code>: <message>")`, or `HTTP <status>: <body>` when the body isn't Dataverse JSON (`dataverseManager.ts:966-980`). **The status number and every response header are lost.** The typed calls behave the same way.
- **Outer non-2xx with a multipart body**: resolved and parsed like a 200 (`resolveMultipartErrors`, `:962-965`).
- **A failed part**: Dataverse answers 200 overall, and the part carries its own 4xx/5xx status with the Dataverse error JSON as `body`. Nothing is thrown, so DRS sees `{ status: 412, body: { error: { code, message } } }`.
- **Continue-on-error and #198:**
  - Without `Prefer: odata.continue-on-error`, Dataverse stops at the first failure, and the host throws `missing operation results` (`dataverseBatch.ts:356-364`).
  - With it, `[bad, good]` fails Content-ID correlation (`:350-353`; DRS `batchExecutor.ts:199-211`).
  - **Neither can happen with one part**: there is one expected id and one result, and the position fallback (`:341-344`) covers a missing Content-ID. Raw transport never needs continue-on-error.
- **Malformed multipart** throws `Malformed Dataverse batch response…` (`:283-331`). It's rare, but it has to be mapped (§2.3).

### 1.6 Header consent

`src/main/managers/dataverseHeaderConsentManager.ts`:

- `authorizeBatch` (`:43-65`) collects the outer headers plus **every header on every part**, scoped as `Operation 1: PATCH /api/data/v9.2/…`. **If the list is empty, no prompt is shown.**
- When there are headers, `ensureConsent` (`:67-88`) skips the prompt if the tool already has consent:
  - **Allow for tool** is persisted in settings (`:71, 94-95`);
  - **Allow for this instance** is kept per tab (`:96-97`).

  Consent is **per tool, not per header**. A user who allowed DRS headers on a typed call won't be asked again by the raw path, and the other way round.
- **Reject** throws `Dataverse additional headers were rejected` (`:99-100`). The prompt times out after 120 s (`:28, 146`). DRS's `CONSENT_DECLINED_RE` already matches both messages.

So a raw request with **no part headers never prompts**. PUT column, single-property DELETE and a plain alternate-key Upsert need none (§3).

### 1.7 Impersonation

There are two kinds, and they behave differently:

- **DRS impersonation** sends `CallerObjectId` / `MSCRMCallerID` as a request header (#109). ~~As a part header, Dataverse applies it to that operation.~~ **Checked live (§6): Dataverse ignores it inside a part.** An unknown `CallerObjectId` in the part still returned 204; on the outer `$batch` it returned 403, and the signed-in user's own id returned 204. So `rawRequest` puts these two headers on the outer request (the third `executeBatch` argument) and everything else in the part. They go through consent there, as on the typed calls.
- **ToolBox impersonation** is the host's own "run as user" for a tool window. The typed handlers wrap their call in `withImpersonation(callerObjectId, …)`, which adds `CallerObjectId` to the HTTP request (`index.ts:2161, 2170`, `dataverseManager.ts:91-93, 910`). **The `executeBatch` / `executeTransaction` handlers don't** (`index.ts:2135-2155`), and the header would only reach the outer `$batch` anyway. A request sent through raw transport while ToolBox impersonation is on runs **as the signed-in user**. `@pptb/types` doesn't expose whether ToolBox impersonation is on, so DRS can't detect it.

### 1.8 Encoding (host-wide)

`makeHttpRequest` builds the response with `data += chunk`, which decodes each network chunk to UTF-8 separately (`dataverseManager.ts:921-930`). A multi-byte character split across two chunks would come out as two replacement characters. This affects every typed call as well as `$batch`, so raw transport doesn't make it worse. It's an upstream candidate (§7). It's also the second reason (after §1.2's string-only body) why binary bodies can't go through the host.

### 1.9 Summary

| Can | Can't |
|---|---|
| Any `GET/POST/PATCH/PUT/DELETE` under `/api/data/v9.2/`, including property paths and alternate-key addresses | Absolute URLs (`@odata.nextLink` / `deltaLink` must be made relative), other API versions, `HEAD`/`OPTIONS` |
| Any request header, sent verbatim on the part | A part header without consent (when there is one) |
| Real per-operation status, all part response headers, parsed body | Outer response headers beyond the three `x-ms-*` (no `Retry-After`) |
| Secondary / numbered connection via `connectionTarget` | ToolBox-level impersonation |
| JSON or string bodies up to ~16 MB | Binary request or response bodies; responses over 16 MB |
| A literal `?` in the query string | — it's truncated; encode it as `%3F` |

---

## 2. Proposed bridge API

### 2.1 `dvHost.rawRequest`

In `src/host/pptbBridge.ts`, next to the existing `executeBatch` wrapper (`:421-434`):

```ts
export interface RawRequest {
  method: 'GET' | 'POST' | 'PATCH' | 'PUT' | 'DELETE';
  /** Path relative to /api/data/v9.2/, e.g. `accounts(<id>)/name`. No scheme/host. */
  path: string;
  /** Sent verbatim on the batch part. Any entry → host consent prompt (once per tool). */
  headers?: Record<string, string>;
  /** Object → JSON; string → as is. No binary. */
  body?: unknown;
  /** Defaults to the primary connection. */
  connectionTarget?: DataverseAPI.ConnectionTarget;
}

export interface RawResponse {
  status: number;
  statusText?: string;
  /** The part's response headers, lowercased (ETag, OData-EntityId, Location, Preference-Applied…). */
  headers: Record<string, string>;
  /** The three outer x-ms-* headers the host passes through, when present. */
  batchHeaders?: Record<string, string>;
  body?: unknown;
}

dvHost.rawRequest = async (req: RawRequest): Promise<RawResponse> => {
  const part: DataverseAPI.BatchRequest = {
    method: req.method,
    url: toBatchUrl(req.path), // see below
    contentId: '1',
    ...(req.headers && Object.keys(req.headers).length ? { headers: req.headers } : {}),
    ...(req.body !== undefined ? { body: req.body } : {}),
  };
  const [r] = await dv().executeBatch([part], req.connectionTarget);
  if (!r) throw new Error('Malformed Dataverse batch response: no operation result');
  return { status: r.status, statusText: r.statusText, headers: r.headers ?? {}, batchHeaders: r.batchHeaders, body: r.body };
};
```

Notes:

- **`toBatchUrl(path)`** is a pure function with unit tests. It:
  - strips a leading `/`, and any absolute `https://<org>/api/data/v9.x/` prefix (the same logic as `executeAbsoluteUrl`, `dataverseExecutor.ts:608-614`, so it moves into one shared helper);
  - percent-encodes every literal `?` after the first one, to get around the §1.2 truncation;
  - throws on an empty path, `..` or `$batch`, so DRS reports a clear "not sent" before the host rejects it with a generic message.
- **`connectionTarget`** is passed explicitly, unlike the existing wrappers, which pass `undefined`. This lets #231 use it.
- ~~No outer headers, ever.~~ No outer headers except impersonation, which only counts there (§1.7).
- `minAPI` is 1.2.8 (#408), so every supported host has `executeBatch`: there is no capability gate. Outside PPTB, `dv()` throws `HostNotAvailableError` as for every other host call.

### 2.2 Which headers go on the part

Inside a part nothing is host-owned, so Execute can finally send **what the Code tab shows**. Add `toRawPartHeaders(composed)` in `requestHeaderPolicy.ts`. It drops:

- `Authorization`, credential headers (`isCredentialHeader`, #217) and `Content-Length`;
- `Accept`, `OData-Version` and `OData-MaxVersion` when they hold the defaults. Dataverse's defaults apply inside a part, and leaving them out avoids a needless prompt.

It keeps:

- `Prefer`, from the mode's own Prefer pane rather than `HOST_OWNED_HEADERS`;
- `If-Match` / `If-None-Match`, except with a request on one column's URL: Dataverse refuses any `If-Match` there with 405 "Versioning is not supported at property level" (`0x80060888`, live). `composeHeaderItems` leaves the precondition out for Update's PUT and Delete's single-property scope, and their Precondition panes say why;
- the `MSCRM.*` bypass headers (not verified live: no plug-in to bypass in the test environment);
- impersonation, which `rawRequest` then moves to the outer request (§1.7);
- `Content-Type`, only when it isn't JSON.

When the result is `{}`, no prompt is shown.

### 2.3 Mapping onto `ExecResult`

A new pure function in `dataverseExecutor.ts`, next to `ok()` / `failure()` (`:273-371`):

```ts
function fromRaw(r: RawResponse, ms: number): ExecResult {
  const headers = r.headers ?? {};
  if (r.status >= 200 && r.status < 300 || r.status === 304) {
    return { ...ok(r.status, r.body ?? null, ms), statusText: r.statusText ?? STATUS_TEXT[r.status] ?? 'OK',
             headers, transport: 'batch', batchHeaders: r.batchHeaders };
  }
  const err = isDataverseErrorBody(r.body) ? r.body.error : { message: typeof r.body === 'string' ? r.body : `HTTP ${r.status}` };
  return { ...failure(r.status, err, ms, r.statusText), headers, transport: 'batch', batchHeaders: r.batchHeaders };
}
```

- **Status:** the real one. `outcomeFor` buckets it as usual.
  - A 304 on a conditional GET counts as success. The typed path has to fake that today (`executeRetrieveSingle`, `:575-577`).
  - A 412 on `If-Match` is a real 412, no longer read out of a message.
- **Headers:** `ExecResult.headers` becomes the part's real headers. The `headers` doc comment (`:121-124`) changes, and two optional fields are added:
  - `transport?: 'batch'`, so the Results → Headers view can say the headers came from the `$batch` part;
  - `batchHeaders?: Record<string, string>`, for the three outer `x-ms-*` headers, shown in a separate group.

  Typed results keep `headers: {}`, so the empty state in `ResultsView.tsx:553` stays true for them.
- **Body / bytes:** unchanged (`jsonBytes`).
- **Errors that never reached a part:** everything `executeBatch` throws goes through the existing `mapError(err, ms, sentHeaders)` (`:298-356`). That covers consent declined → 403 "Headers not approved", header rejected → 400, outer 429 / 401, size limit, and malformed multipart. Two new patterns:
  - `exceeds the maximum size` → a "response too large for the batch path (16 MB)" message that names the typed alternative where there is one;
  - `Malformed Dataverse batch response` → "ToolBox couldn't read Dataverse's answer; the request may have run".
- **Throttling:**
  - A part-level 429 keeps its headers, so `headers['retry-after']` is there if Dataverse sends one on a part.
  - The usual case is the outer 429, which becomes a message-only error. `serviceProtection.ts` still recognises the code, but no duration is available (see #361 in §3).
  - When `batchHeaders` carries the `x-ms-ratelimit-*` pair, the Headers view shows it. That's a debugging aid, not a control signal (per the docs).

---

## 3. What each piece unblocks

| Item | Today (dev @ d415ffa) | With `rawRequest` | Needs consent? |
|---|---|---|---|
| **Update — PUT single column** | Refused with 400 (`dataverseExecutor.ts:775-787`); advisory `UpdateMode.tsx:485`; README 📝 (`:521`) | `PUT <set>(<id>)/<col>` with body `{ "value": … }` → 204. Clearing a value: `DELETE` the property (below) | No, unless the user adds headers |
| **Delete — single property** | Refused with 400 (`:864-875`); advisory `DeleteMode.tsx:393-399` | `DELETE <set>(<id>)/<col>` → 204 | No |
| **Upsert — alternate key** | Refused with 400 (`:816-828`); advisory `UpsertMode.tsx:275`; README 📝 (`:522`) | `PATCH <set>(<key>='v',…)` → 204, or 201/200 with `Prefer: return=representation`. Create-only / update-only via `If-None-Match: *` / `If-Match: *` | Only when Prefer or a precondition is on |
| Upsert / Update by GUID: tell 201 from 204, return the row | Always reported as 204 with no row (`:798-800, 844-845`) | Possible, but **not built**: PR 4 was dropped (see Decisions) | Prefer → yes |
| **Response headers** (#59 follow-up) | Headers tab always empty (`ResultsView.tsx:553`) | Real part headers for every request sent raw (`ETag`, `OData-EntityId`, `Location`, `Preference-Applied`, `x-ms-*`). Enables "use this ETag as If-Match" (#227 can capture it) | — |
| **#361 Retry-After** | No duration (`serviceProtection.ts:5-8`) | **Not unblocked.** A throttled `$batch` is an outer non-multipart 429: the host rejects with the message only, and `Retry-After` isn't in `batchHeaders` (§1.3, §1.5). A part-level 429, if Dataverse ever sends one, would carry it. **Blocked on ToolBox** (§7, ask 1) | — |
| **#233 change tracking** | Blocked "on #109 or #225" | **Doesn't need #225.** #109 shipped. Typed `queryData` takes a custom `Prefer` (host appends it, `dataverseManager.ts:886-895`), returns the full body including `@odata.deltaLink` (`:744-762`), and the delta follow-up is a relative `queryData` call. DRS just has to add `odata.track-changes` to `readPrefer` (`requestHeaderPolicy.ts:141-145`). Raw adds only the `Preference-Applied` check | Prefer → yes (as with page size today) |
| **#231 compare responses** | — | Diffing works on bodies with typed calls. Raw adds status and header diffs. `rawRequest` takes `connectionTarget`, so "run on secondary" works for the raw-only shapes too | — |
| **#267 request tests** | — | Status and header assertions become meaningful for requests sent raw | — |
| **Binary modes' 501s** (`:1439-1449`; README `:141-145`) | `hostNotSupported` | **Partly, and not because of #225.** The message pipelines (`Initialize*BlocksUpload` → `UploadBlock` (base64 `BlockData`) → `Commit*BlocksUpload`; `Initialize*BlocksDownload` → `DownloadBlock` (base64 `Data`); `DeleteFile`) are JSON actions that typed `execute()` can already send. The blocker is DRS not reading file bytes (#222). `GET …/$value`, chunked `PATCH` with `application/octet-stream` + `Content-Range`, and full-size image PATCH need binary bodies, which neither host path supports (§1.2, §1.8). → Separate follow-up issue, not part of #225 | — |

The README Web API coverage rows for single-property PUT/DELETE and alternate-key upsert move from 📝 to ✅ (1.2.8). The binary rows stay 📝.

---

## 4. Risks

1. **16 MB response cap.**
   - Applies to: only the raw-routed shapes, whose responses are tiny (204, or one row).
   - Mitigation: don't route reads or multi-row responses through raw. Map the size error clearly (§2.3).
2. **ToolBox impersonation is silently dropped** (§1.7).
   - Applies to: a user with ToolBox's own "impersonate" turned on who runs one of the three raw shapes. It runs as them, not as the impersonated user. DRS can't detect this.
   - Mitigations:
     - (a) DRS's own `CallerObjectId` header still works, inside the part;
     - (b) the result card for a raw-sent request says "Sent through `$batch`: ToolBox impersonation doesn't apply; use the Impersonation header";
     - (c) upstream ask 2.

   This is also why GUID-addressed writes stay typed by default.
3. **Consent prompts.**
   - The three shapes send no headers by default, so they don't prompt.
   - Any header the user enables (Prefer, preconditions, bypass, impersonation) prompts once per tool, exactly as the typed path does today.
   - A decline already maps to 403 "Headers not approved".
4. **Error parsing.**
   - Single-part batches avoid #198 and the continue-on-error correlation errors entirely (§1.5).
   - Remaining host throws: outer status errors, size, malformed multipart, consent. They all go through `mapError`.
   - Part errors arrive as data (status + Dataverse error body), which is *better* than the typed path's message parsing.
5. **Behaviour differences from the typed calls:**
   - The default part `Content-Type` is `application/json; type=entry`, even for a PUT of a property value. Dataverse accepts it (live, §6).
   - No `odata.include-annotations` unless DRS sends `Prefer`. This only matters for reads, which stay typed.
   - No host `return=representation`: DRS sends `Prefer` only when the mode asks.
   - The URL must be relative, and a second literal `?` is truncated. `toBatchUrl` encodes it, and alternate-key values containing `?` need this too.
   - Entitlement and service-protection accounting: a single-part `$batch` costs one request plus one operation. That's no worse in practice. The docs say batching doesn't change entitlement counting.
6. **Code tab drift.**
   - The Code tab keeps showing the *direct* request (`PUT accounts(<id>)/name`), not the `$batch` envelope, because that's what the user authored and what they'd run elsewhere.
   - The result says how it was sent.
   - The PPTB code target, if it emits `dataverseAPI` calls, should emit the `executeBatch` form for these three shapes.
7. **Host-wide UTF-8 chunk decoding** (§1.8): pre-existing, not specific to raw transport. Track upstream.

---

## 5. Implementation plan

Every PR branches off `dev`, targets `dev`, and is merged with a merge commit by the owner. Each one carries a CHANGELOG entry (except this doc).

### PR 1 — bridge + mapping (no user-visible change)

- `pptbBridge.ts`:
  - `RawRequest` / `RawResponse`;
  - `dvHost.rawRequest`;
  - `toBatchUrl` (move the absolute-URL stripping from `executeAbsoluteUrl` into it and reuse it there).
- `requestHeaderPolicy.ts`: `toRawPartHeaders`.
- `dataverseExecutor.ts`:
  - `fromRaw`;
  - `ExecResult.transport` / `batchHeaders`;
  - two new `mapError` patterns (size, malformed).
- Tests (vitest, pure):
  - `toBatchUrl`: relative and absolute input, leading `/`, a second `?` encoded, `..` / `$batch` rejected;
  - `toRawPartHeaders`: credential and default headers dropped, Prefer kept, `{}` when only defaults are left;
  - `fromRaw`: 204 / 201 + body / 304 / 412 with the Dataverse error body / part-level 429 keeping `retry-after` / string body;
  - `rawRequest` with a mocked `window.dataverseAPI.executeBatch`: one part, no `headers` key when there are none, `connectionTarget` passed through, empty result → error;
  - `mapError` for the IPC-wrapped size and malformed messages.

### PR 2 — the three refused shapes

- `executeUpdate` (PUT), `executeDelete` (single-property) and `executeUpsert` (alternate key) call `rawRequest`. Today's 400 refusals are removed (no fallback: see Decisions).
- Drop or replace the advisories in `UpdateMode` / `DeleteMode` / `UpsertMode` with an info note: "Sent through ToolBox's `$batch`; ToolBox impersonation doesn't apply."
- `WriteResultCard` copy for raw-sent results.
- README: coverage table (📝 → ✅ 1.2.8), Update/Upsert/Delete bullets (`:58, :63, :65`).
- Tests:
  - executor tests per shape (URL, method, body, headers or none, result mapping);
  - mode render tests: the advisory is gone and Execute is enabled;
  - `executeCodeParity`: the executed request equals the Code tab request apart from the envelope.
- **Live E2E in PPTB** (owner's protocol, test records only, cleaned up): run each shape, plus the checks in §6.

### PR 3 — Headers view (done)

- `ResultsView` Headers tab: render `result.headers` when `transport === 'batch'`, with the three `batchHeaders` in their own group. Keep the empty-state text for typed results.
- "Use as If-Match" action on an `etag` header. Fall back to the body's `@odata.etag` when the header is missing.
- Tests: render with and without headers, the ETag action fills the precondition.
- **As built:** `ResultsView` takes `onUseEtag`. Only Upsert passes it: of the three shapes, only an alternate-key Upsert returns an ETag (PUT and clearing a column answer 204 with none, and take no precondition, §2.2). A typed result has no `etag` header and no row, so the body fallback doesn't come into play there today. The batch-headers group is labelled "From the $batch response" and says the rate-limit counts are for reading, not pacing.

### ~~PR 4~~ — dropped (owner, 2026-11-22): "Send through `$batch` to see headers"

- A per-request toggle (off by default) that sends GUID-addressed Create / Update / Upsert / Delete and Retrieve Single through `rawRequest`. That gives the real status (201 vs 204), the returned row, and headers.
- Shown with the impersonation and 16 MB caveats.
- Tests as for PR 2.

### Re-scoping of dependants (issue edits for the owner)

- **#361:** blocked on the ToolBox change in §7 ask 1, not on #225. Suggest moving it out of v1.8.0 unless ToolBox ships the ask.
- **#233:** remove the #225 dependency; it can go ahead on typed `queryData`.
- **#231:** independent; it uses `connectionTarget` for the raw shapes once PR 1 lands.
- **New issue:** message-based binary uploads/downloads (read bytes, base64 blocks through typed `execute()`). Depends on #222, not #225.

---

## 6. Live checks for PR 2 (not verifiable from source)

1. A `PUT accounts(<id>)/name` part with the host's default `Content-Type: application/json; type=entry` → 204. If Dataverse rejects `type=entry`, send `Content-Type: application/json` explicitly, which costs one consent prompt per tool.
2. Part response headers actually returned by Dataverse for: PUT property, DELETE property, alternate-key PATCH (with and without `return=representation`), GET single (is `ETag` a part header or only `@odata.etag`?).
3. Alternate-key values with `'`, `?`, `#`, `&` and non-ASCII, after `toBatchUrl` encoding.
4. `CallerObjectId` as a part header → `createdonbehalfby` / `modifiedonbehalfby` show the real caller.
5. A part-level failure (412 on `If-Match` with a stale ETag) → `{ status: 412, body.error }`, nothing thrown.
6. With ToolBox impersonation on: confirm the raw request runs as the signed-in user, to word the note accurately.

**Results (2026-11-22, PPTB with `minAPI` 1.2.8, test rows created and deleted):**

1. ✅ `PUT accounts(<id>)/name` with the default part `Content-Type` → 204, value written. **But** with `If-Match: *` (Update's default precondition) → 405 `0x80060888` "Versioning is not supported at property level". The same for `DELETE accounts(<id>)/telephone1` with an ETag. Without a precondition, the DELETE → 204 and the column is null. Fixed: no precondition with a request on one column's URL (§2.2).
2. ✅ Part headers come back: PUT / DELETE → `odata-version`, `x-content-type-options`; alternate-key PATCH with `return=representation` → `etag`, `preference-applied`, `content-type`, and the row. GET single stays typed.
3. ✅ `msfp_unsubscribedrecipients(msfp_emailaddress='drs-e2e-225+o''q?x#y&zé@example.invalid')` → 201 Created, then 200 on the same key; the stored value round-trips exactly. ❌ With a `%` in the value (`…z%é…`), ToolBox refused the request before sending it ("URL contains invalid encoding", §1.2). Fixed: `toBatchUrl` refuses such a path with a message saying why and suggesting the GUID.
4. ❌ `CallerObjectId` in the part is ignored (an unknown id → 204). On the outer request it applies (unknown → 403; own id → 204). Fixed: `rawRequest` sends impersonation outer (§1.7).
5. ✅ Alternate-key PATCH with a stale `If-Match` → `{ status: 412, body.error.code: 0x80060882 }`, shown as a result, nothing thrown.
6. Not run: ToolBox impersonation wasn't on. The source (§1.7) shows the `executeBatch` handler doesn't apply it, so the note stands.

---

## 7. Upstream asks (desktop-app) — for the owner to file

1. **Throttling:** add `retry-after` to `DATAVERSE_BATCH_RESPONSE_HEADERS` (`dataverseBatch.ts:13`). On a non-multipart outer error, reject with an error that carries `status` and the response headers (at least `retry-after`), not only `<code>: <message>` (`dataverseManager.ts:966-980`). Doing the same for typed calls would unblock #361 for every mode.
2. **Impersonation:** wrap `EXECUTE_BATCH` / `EXECUTE_TRANSACTION` in `withImpersonation`: an outer `CallerObjectId` applies to every part (live, §6). Or document that batch ignores ToolBox impersonation (`index.ts:2135-2155`).
3. **Query truncation:** `split("?", 2)` drops everything after a second `?` (`dataverseBatch.ts:115`). Use `indexOf("?")`.
3a. **`%` in a path:** `validateDataverseRelativeUrl` throws "URL contains invalid encoding" when a round of its traversal check's decoding throws, so a value with a literal `%` (`%25`) followed by anything that isn't a valid escape once decoded can't be sent (`dataverseBatch.ts:116-125`). Treat a decoding error as "no further decoding" (an undecodable string can't hide a traversal), or decode only the `%2E` / `%2F` / `%5C` forms it looks for.
4. **UTF-8 decoding:** collect `Buffer` chunks and decode once (`dataverseManager.ts:921-930`). This also opens the door to binary responses.
5. (Nice to have) a first-class single-request API (`dataverseAPI.request({ method, url, headers, body })` returning `{ status, headers, body }`) with the typed calls' 128 MB limit, binary bodies and impersonation. That would make this whole workaround unnecessary.
