// Code generation for the Code / Editor tabs.
//
// Each generator returns a ready-to-paste snippet. Conventions:
//   - Always print a leading comment explaining context (auth, where to run)
//   - Always wrap async work in an IIFE so the snippet is portable across
//     classic <script> / Node REPL / Power Apps event handlers
//   - Always show error handling — silent failures are the #1 footgun
//   - Preserve the user's actual headers (Prefer, MSCRMCallerID, etc.)

import type { BuiltRequest } from './urlBuilder';
import { ENV } from '../mock/environment';

export type CodeFormat = 'fetch' | 'xrm' | 'xrm-batch' | 'xhr' | 'powerautomate' | 'csharp' | 'powershell' | 'curl' | 'json';

export const FORMAT_LABELS: Record<CodeFormat, string> = {
  fetch:        'fetch',
  xrm:          'Xrm.WebApi',
  'xrm-batch':  'Xrm.WebApi.executeMultiple',
  xhr:          'XMLHttpRequest',
  powerautomate:'Power Automate',
  csharp:       'C#',
  powershell:   'PowerShell',
  curl:         'cURL',
  json:         'JSON',
};

export const FORMAT_LANG: Record<CodeFormat, string> = {
  fetch: 'javascript', xrm: 'javascript', 'xrm-batch': 'javascript', xhr: 'javascript',
  // Power Automate is rendered as a "Field  Value" cheatsheet you can paste
  // into the Dataverse "List rows" / "Get a row by ID" action — not JSON.
  powerautomate: 'plaintext',
  csharp: 'csharp', powershell: 'powershell',
  curl: 'shell', json: 'json',
};

export interface CodegenInputs {
  method: string;
  built: BuiltRequest;
  /** Final header map (already includes the Prefer header composed from PreferEditor) */
  headers: Record<string, string>;
  /** True for Retrieve NextLink — generators use rawNextLink directly instead of building a URL */
  isNextLink?: boolean;
  rawNextLink?: string;
  /**
   * JSON body for write requests (Create / Update / Upsert). When present,
   * generators serialize it inline instead of emitting a placeholder.
   * The shape already includes @odata.bind keys for lookups and the comma-
   * separated string form for multi-select picklists — caller has done the
   * encoding via engine/urlBuilder.buildCreateBody.
   */
  body?: Record<string, unknown>;
  /**
   * Logical-name (NOT entity-set-name) of the target entity for Xrm.WebApi
   * .createRecord — emitted as the first argument. Optional because it's
   * already exposed via `built.entityLogical`; carried separately when callers
   * want to override (e.g. polymorphic Create-from).
   */
  entityLogical?: string;
  /**
   * Multi-request batch — used by Associate (POST per target) and Disassociate
   * (DELETE per related id) on collection-valued nav props. Per docs, each
   * association/disassociation is a separate HTTP call to the same/different
   * URLs:
   *
   *   POST /accounts(<id>)/contact_customer_accounts/$ref  body: {@odata.id: ".../contacts(t1)"}
   *   POST /accounts(<id>)/contact_customer_accounts/$ref  body: {@odata.id: ".../contacts(t2)"}
   *
   * When set + length > 1, generators emit a loop / sequential calls instead
   * of a single request. `built.relativeUrl` + `body` still drive the URL bar
   * and single-request preview.
   */
  multiRequests?: MultiRequest[];
}

export interface MultiRequest {
  method: string;
  /** Path + query (no host) — same form as built.relativeUrl */
  relativeUrl: string;
  body?: Record<string, unknown>;
  /** Inline label for the request (e.g. "John Doe"). Optional. */
  description?: string;
}

// ── Power Automate field-spec (consumed by PowerAutomatePane) ──
export interface PowerAutomateField {
  /** Label that mirrors the Power Automate connector UI verbatim */
  label: string;
  /** Value ready to paste into that field */
  value: string;
  /** Inline hint shown under the input (e.g. "logical names, comma-separated") */
  hint?: string;
  /** Multi-line shown via Textarea instead of Input (Filter Rows can be long) */
  multiline?: boolean;
}

export interface PowerAutomateActionSpec {
  /** Display title — e.g. "List rows" */
  actionName: string;
  /** Connector context line — e.g. "Microsoft Dataverse connector" */
  connector: string;
  /** Banner shown above the fields */
  banner?: string;
  fields: PowerAutomateField[];
  /** Bulleted notes under the fields */
  notes?: string[];
  /** Fallback HTTP-action shape shown only when the high-level action can't express the request */
  httpFallback?: {
    method: string;
    uri: string;
    note: string;
  };
}

export function generateCode(fmt: CodeFormat, i: CodegenInputs): string {
  switch (fmt) {
    case 'fetch':         return genFetch(i);
    case 'xrm':           return genXrm(i);
    case 'xrm-batch':     return genXrmBatch(i);
    case 'xhr':           return genXhr(i);
    case 'powerautomate': return genPowerAutomate(i);
    case 'csharp':        return genCsharp(i);
    case 'powershell':    return genPowershell(i);
    case 'curl':          return genCurl(i);
    case 'json':          return genJson(i);
  }
}

// ────────────────────────────────────────────────────────────
// Shared helpers
// ────────────────────────────────────────────────────────────
function fullUrl(i: CodegenInputs): string {
  if (i.isNextLink && i.rawNextLink) return i.rawNextLink;
  return `https://${ENV.host}${i.built.relativeUrl}`;
}

/** Strip headers a generator already inserts elsewhere (e.g. Authorization). */
function stripAuth(headers: Record<string, string>): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [k, v] of Object.entries(headers)) {
    if (!v) continue;
    if (k.toLowerCase() === 'authorization') continue;
    out[k] = v;
  }
  return out;
}

/** Escape a string for inclusion inside a single-quoted JS literal. */
function jsEscape(s: string): string {
  return s.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
}

/** Render headers as a JS object body, indented. */
function jsHeaderObject(headers: Record<string, string>, indent: string): string {
  const entries = Object.entries(headers).filter(([, v]) => v);
  if (entries.length === 0) return '';
  return entries.map(([k, v]) => `${indent}'${jsEscape(k)}': '${jsEscape(v)}'`).join(',\n');
}

/**
 * Pretty-print a value as a JavaScript object literal — preserves @odata.bind
 * keys (which contain `@`) by quoting them, formats nested values with stable
 * indentation, and keeps numbers/booleans/null unquoted.
 *
 * Why not JSON.stringify? Generated code reads better with single quotes,
 * trailing commas, and a leading indent on continuation lines. JSON output
 * also unconditionally quotes every key, which makes lookup binds noisier.
 */
function formatJsObject(obj: Record<string, unknown>, baseIndent: string): string {
  const keys = Object.keys(obj);
  if (keys.length === 0) return '{}';
  const nextIndent = baseIndent + '  ';
  const lines = keys.map(k => {
    const v = obj[k];
    const keyOut = /^[a-zA-Z_$][a-zA-Z0-9_$]*$/.test(k) ? k : `'${jsEscape(k)}'`;
    return `${nextIndent}${keyOut}: ${formatJsValue(v, nextIndent)}`;
  });
  return `{\n${lines.join(',\n')},\n${baseIndent}}`;
}

/**
 * Decode a single function alias value (from queryParts) back to a JS-literal-
 * compatible form. OData inline literals follow predictable rules:
 *   • 'value' → 'value' (single-quoted string)
 *   • bare number → number
 *   • bare true/false → boolean
 *   • bare guid → string literal (we wrap for safety)
 * Used by genXrm to surface function params inside the request descriptor.
 */
function decodeFunctionAlias(odataLiteral: string): string {
  const trimmed = odataLiteral.trim();
  // Single-quoted string — keep as JS string literal
  if (trimmed.startsWith("'") && trimmed.endsWith("'")) {
    const inner = trimmed.slice(1, -1).replace(/''/g, "'");
    return `'${inner.replace(/'/g, "\\'")}'`;
  }
  // Booleans
  if (trimmed === 'true' || trimmed === 'false') return trimmed;
  // Number
  if (/^-?\d+(?:\.\d+)?$/.test(trimmed)) return trimmed;
  // GUID or unrecognized — pass through as quoted string for safety
  return `'${trimmed.replace(/'/g, "\\'")}'`;
}

function formatJsValue(v: unknown, indent: string): string {
  if (v === null) return 'null';
  if (typeof v === 'string')  return `'${jsEscape(v)}'`;
  if (typeof v === 'number' || typeof v === 'boolean') return String(v);
  if (Array.isArray(v)) {
    if (v.length === 0) return '[]';
    return `[ ${v.map(x => formatJsValue(x, indent)).join(', ')} ]`;
  }
  if (typeof v === 'object') {
    return formatJsObject(v as Record<string, unknown>, indent);
  }
  return JSON.stringify(v);
}

// ────────────────────────────────────────────────────────────
// fetch — modern browser / Node 18+ / Power Apps form context
// ────────────────────────────────────────────────────────────
function genFetch(i: CodegenInputs): string {
  // Multi-request path — Associate / Disassociate over a collection sends N
  // requests sequentially. Per docs, each association is its own call to a
  // $ref endpoint, so we emit a loop with one body per target.
  if (i.multiRequests && i.multiRequests.length > 1) {
    return genFetchMulti(i);
  }
  const url = fullUrl(i);
  const hdrs = stripAuth(i.headers);
  const headerBlock = jsHeaderObject(hdrs, '      ');
  const isWrite = i.method !== 'GET';
  const bodyLine = isWrite
    ? `\n      body: JSON.stringify(payload),`
    : '';
  const payloadDecl = isWrite
    ? `  const payload = ${formatJsObject(i.body ?? {}, '  ')};\n\n`
    : '';

  return `// Dataverse Web API — fetch
// Run from a browser context (or Node 18+). Provide a bearer token via the
// Authorization header below.  In Power Apps / D365 form scripts you can also
// call Xrm.WebApi directly — see the Xrm.WebApi tab.

(async () => {
${payloadDecl}  const url = '${url}';

  const res = await fetch(url, {
    method: '${i.method}',
    headers: {
      'Authorization': 'Bearer <access-token>',
${headerBlock}
    },${bodyLine}
  });

  if (!res.ok) {
    const err = await res.json().catch(() => null);
    throw new Error(\`Dataverse \${res.status}: \${err?.error?.message ?? res.statusText}\`);
  }

  // 204 No Content (e.g. Delete) doesn't carry a JSON body
  const data = res.status === 204 ? null : await res.json();
  console.log('Result:', data);
  if (data && Array.isArray(data.value)) {
    console.log(\`Returned \${data.value.length} record(s)\`);
    if (typeof data['@odata.count'] === 'number') {
      console.log(\`@odata.count: \${data['@odata.count']}\`);
    }
    if (data['@odata.nextLink']) {
      console.log('nextLink:', data['@odata.nextLink']);
    }
  }
})().catch(console.error);
`;
}

/**
 * Multi-request fetch — Associate / Disassociate sends N sequential calls.
 * Per docs, each $ref operation is its own request — no batch implied.
 * Generated code stages every request in an array and loops with await so
 * errors surface in order; for high counts the user should switch to a
 * $batch request (separate mode, lands later).
 */
function genFetchMulti(i: CodegenInputs): string {
  const hdrs = stripAuth(i.headers);
  const headerBlock = jsHeaderObject(hdrs, '      ');
  const reqs = i.multiRequests ?? [];
  const requestArray = reqs.map(r => {
    const url = `https://${ENV.host}${r.relativeUrl}`;
    const bodyLine = r.body ? `, body: JSON.stringify(${formatJsObject(r.body, '    ')})` : '';
    const desc = r.description ? `  // ${r.description}\n` : '';
    return `${desc}  { method: '${r.method}', url: '${url}'${bodyLine} },`;
  }).join('\n');

  return `// Dataverse Web API — fetch (multi-request)
//
// Each association / disassociation is its own HTTP call per the docs:
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/associate-disassociate-entities-using-web-api
//
// For more than a handful of targets, prefer a \$batch request — sending
// many small requests in sequence is slower and not transactional.

(async () => {
  const requests = [
${requestArray}
  ];

  const sharedHeaders = {
    'Authorization': 'Bearer <access-token>',
${headerBlock}
  };

  const results = [];
  for (const r of requests) {
    const res = await fetch(r.url, {
      method: r.method,
      headers: sharedHeaders,
      ${reqs.some(r => r.body) ? 'body: r.body,' : ''}
    });
    if (!res.ok) {
      const err = await res.json().catch(() => null);
      console.error(\`\${r.method} \${r.url} → \${res.status}\`, err?.error?.message ?? res.statusText);
      results.push({ status: res.status, ok: false });
      continue;
    }
    results.push({ status: res.status, ok: true });
  }

  console.log(\`\${results.filter(r => r.ok).length}/\${results.length} requests succeeded\`);
})().catch(console.error);
`;
}

// ────────────────────────────────────────────────────────────
// Xrm.WebApi — model-driven app / D365 form scripts
// ────────────────────────────────────────────────────────────
function genXrm(i: CodegenInputs): string {
  if (i.isNextLink && i.rawNextLink) {
    return `// Xrm.WebApi doesn't take a raw nextLink URL — follow it via fetch:
// (Run this in a model-driven app form script context)

const nextLink = '${i.rawNextLink}';
const res = await fetch(nextLink, {
  headers: { 'Accept': 'application/json' },
});
const page = await res.json();
console.log('Next page:', page.value);
`;
  }
  // Multi-request — Xrm.WebApi doesn't expose Associate/Disassociate $ref
  // endpoints directly; the documented path is Xrm.WebApi.online.execute() with
  // a request descriptor OR fall back to fetch. We surface the fetch loop.
  if (i.multiRequests && i.multiRequests.length > 1) {
    return `// Xrm.WebApi doesn't expose Associate / Disassociate \$ref endpoints
// directly — the documented path is Xrm.WebApi.online.execute() with a
// request descriptor, or fall back to a fetch loop. We emit the fetch loop:

${genFetchMulti(i)}`;
  }

  // No 'account' hardcoded fallback — empty string is honest; downstream
  // branches handle the no-entity case (unbound actions / functions) by
  // omitting the entity/collection field rather than fabricating one.
  const logical = i.built.entityLogical;
  const recordId = i.built.recordId;
  // Re-derive the options string from queryParts (decoded for readability inside Xrm.WebApi)
  const optsBody = i.built.queryParts
    .map(p => `${p.key}=${decodeURIComponent(p.value)}`)
    .join('&');
  const options = optsBody ? `?${optsBody}` : '';

  // Prefer: odata.maxpagesize → maxPageSize parameter on retrieveMultipleRecords
  let maxPageSize: number | null = null;
  const prefer = i.headers['Prefer'];
  if (prefer) {
    const m = prefer.match(/odata\.maxpagesize\s*=\s*(\d+)/i);
    if (m) maxPageSize = Number(m[1]);
  }

  // Detect actions/functions/$ref/workflow paths — these need
  // Xrm.WebApi.online.execute({request}) rather than the high-level helpers.
  const url = i.built.relativeUrl;
  const isAction =
    i.method === 'POST' && (
      url.includes('/Microsoft.Dynamics.CRM.') ||
      url.endsWith('/Merge') ||
      (!recordId && !i.built.entitySet && !!i.body)   // unbound action like POST /WhoAmI(…) is GET so this won't match
    );
  const isFunction = i.method === 'GET' && !recordId && !i.built.entitySet;
  const isRef = url.endsWith('/$ref') || /\$ref\?/.test(url);
  const isWorkflow = url.includes('/workflows(');
  const isRetrieveMultiple = i.method === 'GET' && !recordId && !isFunction;
  const isRetrieveSingle = i.method === 'GET' && !!recordId && !isFunction;
  const isWrite = i.method !== 'GET' && !isAction && !isRef && !isWorkflow;

  // ── Actions / Workflow ─────────────────────────────────────────
  // Xrm.WebApi.online.execute() takes a request descriptor object whose
  // getMetadata() declares the operation type + bound entity. Cleanest path
  // for OOB actions, custom APIs, custom actions, and ExecuteWorkflow.
  if (isAction || isWorkflow) {
    const actionMatch = url.match(/\/Microsoft\.Dynamics\.CRM\.([^/?(]+)/);
    const actionName = actionMatch ? actionMatch[1] : isWorkflow ? 'ExecuteWorkflow' : url.split('/').pop() ?? '';
    const isBound = !!recordId;
    return `// Xrm.WebApi.online.execute — actions / workflows / custom APIs.
// Run inside a model-driven app or D365 form script. The request descriptor
// must implement getMetadata() so Xrm knows which operation to invoke.

(async () => {
  const request = {
${Object.entries(i.body ?? {}).map(([k, v]) => `    ${/^[a-zA-Z_$][\w$]*$/.test(k) ? k : `'${jsEscape(k)}'`}: ${formatJsValue(v, '    ')},`).join('\n')}
${isBound ? `    entity: { entityType: '${logical}', id: '${recordId}' },\n` : ''}    getMetadata: () => ({
      boundParameter: ${isBound ? `'entity'` : 'null'},
      operationType: 0,  // 0 = Action · 1 = Function · 2 = CRUD
      operationName: '${actionName}',
      parameterTypes: { /* parameter type descriptors per CSDL */ },
    }),
  };

  const response = await Xrm.WebApi.online.execute(request);
  if (response.ok) {
    const result = response.status === 204 ? null : await response.json();
    console.log('Action result:', result);
  } else {
    console.error('Action failed:', response.status, response.statusText);
  }
})().catch(err => console.error('Xrm.WebApi.online.execute error:', err));
`;
  }

  // ── Functions ─────────────────────────────────────────────────
  if (isFunction) {
    const fnMatch = url.match(/\/([^/?(]+)(?=\(|$)/);
    const fnName = fnMatch ? fnMatch[1] : url.split('/').pop() ?? '';
    return `// Xrm.WebApi.online.execute — OData function call.
// Functions are GET requests; the descriptor sets operationType: 1.

(async () => {
  const request = {
${Object.entries(i.built.queryParts.filter(p => p.key.startsWith('@'))
        .reduce<Record<string, string>>((acc, p) => { acc[p.key.replace(/^@/, '')] = p.value; return acc; }, {}))
        .map(([k, v]) => `    ${k}: ${decodeFunctionAlias(v)},`).join('\n')}
    getMetadata: () => ({
      boundParameter: null,
      operationType: 1,  // 1 = Function
      operationName: '${fnName}',
      parameterTypes: { /* parameter type descriptors per CSDL */ },
    }),
  };

  const response = await Xrm.WebApi.online.execute(request);
  if (response.ok) {
    const result = await response.json();
    console.log('Function result:', result);
  } else {
    console.error('Function failed:', response.status, response.statusText);
  }
})().catch(err => console.error('Xrm.WebApi.online.execute error:', err));
`;
  }

  // ── Associate / Disassociate ($ref) ────────────────────────────
  // No Xrm.WebApi helper; fall through to a fetch loop. Documented in the
  // multi-request comment block.
  if (isRef) {
    return `// Xrm.WebApi has no direct support for \$ref endpoints (Associate /
// Disassociate). Use the documented Xrm.WebApi.online.execute() pattern
// with a CreateAssociation / DisassociateRecords descriptor, OR fall back
// to a fetch loop — see the fetch tab for the multi-request version.

${genFetchMulti(i)}`;
  }

  if (isRetrieveMultiple) {
    return `// Xrm.WebApi.retrieveMultipleRecords — run inside a model-driven app / D365 form script.
// The options string mirrors the query string of the underlying REST URL.

(async () => {
  const entityLogicalName = '${logical}';
  const options = '${options}';${maxPageSize ? `\n  const maxPageSize = ${maxPageSize};` : ''}

  const result = await Xrm.WebApi.retrieveMultipleRecords(entityLogicalName, options${maxPageSize ? ', maxPageSize' : ''});

  console.log(\`Returned \${result.entities.length} record(s)\`);
  console.log(result.entities);

  // Annotations attach to each entity row as <prop>@OData.Community.Display.V1.FormattedValue, etc.
  // If a @odata.nextLink is returned, fetch it directly to get the next page:
  if (result.nextLink) {
    console.log('Next page:', result.nextLink);
  }
})().catch(err => console.error('Xrm.WebApi error:', err));
`;
  }

  if (isRetrieveSingle) {
    return `// Xrm.WebApi.retrieveRecord — run inside a model-driven app / D365 form script.

(async () => {
  const entityLogicalName = '${logical}';
  const id = '${recordId ?? '<id>'}';
  const options = '${options}';

  const record = await Xrm.WebApi.retrieveRecord(entityLogicalName, id, options);
  console.log('Got record:', record);
})().catch(err => console.error('Xrm.WebApi error:', err));
`;
  }

  // Write group — Create / Update / Delete / Upsert
  if (isWrite) {
    const verb =
      i.method === 'POST'   ? 'createRecord' :
      i.method === 'PATCH'  ? 'updateRecord' :
      i.method === 'DELETE' ? 'deleteRecord' :
      'execute';
    const dataLiteral = verb === 'deleteRecord'
      ? '' // delete doesn't take a body
      : `  const data = ${formatJsObject(i.body ?? {}, '  ')};\n\n`;
    const callArgs = verb === 'deleteRecord'
      ? 'entityLogicalName, id'
      : `entityLogicalName${recordId ? ', id' : ''}, data`;
    return `// Xrm.WebApi.${verb} — write-side request.
// Run inside a model-driven app / D365 form script.
//
// Notes:
//   • @odata.bind keys in \`data\` set lookup values — the value is the OData
//     entity-set URL fragment, e.g. '/contacts(<guid>)'.
//   • Multi-select choice columns are sent as comma-separated integer strings.
//   • return=representation is automatic with Xrm.WebApi.${verb}: the resolved
//     result already carries the freshly-saved row (with system-set values).

(async () => {
  const entityLogicalName = '${logical}';
${recordId ? `  const id = '${recordId}';\n` : ''}${dataLiteral}  const result = await Xrm.WebApi.${verb}(${callArgs});
  console.log('Result:', result);
})().catch(err => console.error('Xrm.WebApi error:', err));
`;
  }
  return '';
}

// ────────────────────────────────────────────────────────────
// Xrm.WebApi.online.executeMultiple — batch / transaction wrapper
// ────────────────────────────────────────────────────────────
//
// Doc: https://learn.microsoft.com/en-us/power-apps/developer/model-driven-apps/clientapi/reference/xrm-webapi/online/executemultiple
//
// `executeMultiple` takes an array of request descriptors. Each descriptor
// has a `getMetadata()` returning `{ boundParameter, operationType,
// operationName, parameterTypes }`. operationType:
//   0 = Action      (custom actions, OOB actions)
//   1 = Function    (OData functions like WhoAmI, FormatAddress)
//   2 = CRUD        (Create / Retrieve / Update / Delete / Associate / etc.)
//
// Use case: batching multiple operations as either a "Change Set" (all-or-
// none transactional) or "Faulted" (continue on error). Even single
// operations wrap fine — the array just has one entry. We emit a batch
// regardless because that's the entire point of this generator.

function genXrmBatch(i: CodegenInputs): string {
  if (i.isNextLink && i.rawNextLink) {
    return `// executeMultiple doesn't follow raw @odata.nextLink URLs — paging on a
// previously-returned cursor lives outside the batch pattern. Use the
// Xrm.WebApi tab (which falls back to fetch for next-link follow-ups) or
// call retrieveMultipleRecords directly with the same options string.
`;
  }

  // No `|| 'account'` fallback — that produced misleading `collection: 'accounts'`
  // output for unbound actions that have no entityLogical. Empty string is
  // honest; downstream code handles the no-entity case by NOT emitting the
  // collection/entity line.
  const logical = i.built.entityLogical;
  const recordId = i.built.recordId;
  const url = i.built.relativeUrl;

  // Match the action detection used by the other generators (genXrm,
  // genFetch's classification, etc.): unbound actions don't have the
  // `/Microsoft.Dynamics.CRM.` namespace prefix on their URL, so we also
  // detect via the structural signature
  // `POST + no recordId + no entitySet + has body` — this catches
  // GrantAccess, Merge, etc. that otherwise fell into the CRUD-write
  // branch and were emitted as `Create` against `accounts`.
  const isAction =
    i.method === 'POST' && (
      url.includes('/Microsoft.Dynamics.CRM.') ||
      url.endsWith('/Merge') ||
      (!recordId && !i.built.entitySet && !!i.body)
    );
  const isFunction = i.method === 'GET' && !recordId && !i.built.entitySet;
  const isWorkflow = url.includes('/workflows(');
  const isRef = url.endsWith('/$ref') || /\$ref\?/.test(url);
  const isRetrieveMultiple = i.method === 'GET' && !recordId && !isFunction;
  const isRetrieveSingle = i.method === 'GET' && !!recordId && !isFunction;
  const isWrite = i.method !== 'GET' && !isAction && !isRef && !isWorkflow;

  // Re-derive the OData query options string from the request's
  // queryParts (decoded for readability inside the JS snippet).
  const optsBody = i.built.queryParts
    .map(p => `${p.key}=${decodeURIComponent(p.value)}`)
    .join('&');
  const options = optsBody ? `?${optsBody}` : '';

  // ── Action / Workflow ─────────────────────────────────────────
  if (isAction || isWorkflow) {
    const actionMatch = url.match(/\/Microsoft\.Dynamics\.CRM\.([^/?(]+)/);
    const actionName = actionMatch ? actionMatch[1] : isWorkflow ? 'ExecuteWorkflow' : (url.split('/').pop() ?? '');
    const isBound = !!recordId;
    const bodyEntries = Object.entries(i.body ?? {});
    return `// Xrm.WebApi.online.executeMultiple — batch one action (extend the array
// to chain more requests). Set the second argument to false for "fail fast"
// (Change Set semantics — first failure rolls back the rest); true for
// "continue on error" (Faulted set semantics).
//
// Doc: https://learn.microsoft.com/en-us/power-apps/developer/model-driven-apps/clientapi/reference/xrm-webapi/online/executemultiple

(async () => {
  const requests = [
    {
${bodyEntries.map(([k, v]) => `      ${/^[a-zA-Z_$][\w$]*$/.test(k) ? k : `'${jsEscape(k)}'`}: ${formatJsValue(v, '      ')},`).join('\n')}
${isBound ? `      entity: { entityType: '${logical}', id: '${recordId}' },\n` : ''}      getMetadata: () => ({
        boundParameter: ${isBound ? `'entity'` : 'null'},
        operationType: 0,  // 0 = Action
        operationName: '${actionName}',
        parameterTypes: { /* parameter type descriptors per CSDL */ },
      }),
    },
  ];

  const results = await Xrm.WebApi.online.executeMultiple(requests);
  results.forEach((r, idx) => {
    console.log(\`Request \${idx} → status \${r.status} \${r.statusText}\`);
  });
})().catch(err => console.error('executeMultiple error:', err));
`;
  }

  // ── Function (operationType: 1) ───────────────────────────────
  if (isFunction) {
    const fnMatch = url.match(/\/([^/?(]+)(?=\(|$)/);
    const fnName = fnMatch ? fnMatch[1] : (url.split('/').pop() ?? '');
    return `// Xrm.WebApi.online.executeMultiple — batched function call.
// Functions use operationType: 1 in the metadata descriptor.

(async () => {
  const requests = [
    {
      getMetadata: () => ({
        boundParameter: null,
        operationType: 1,  // 1 = Function
        operationName: '${fnName}',
        parameterTypes: { /* parameter type descriptors per CSDL */ },
      }),
    },
  ];

  const results = await Xrm.WebApi.online.executeMultiple(requests);
  for (const r of results) {
    if (r.ok) {
      const data = await r.json();
      console.log('Function result:', data);
    } else {
      console.error('Function failed:', r.status, r.statusText);
    }
  }
})().catch(err => console.error('executeMultiple error:', err));
`;
  }

  // ── Associate / Disassociate ($ref) ────────────────────────────
  // For Associate: operationName "Associate", target + relationship name +
  // related entity reference. Disassociate is symmetric. If the mode
  // supplies multiple targets (multiRequests), we emit one descriptor per.
  if (isRef) {
    const isAssociate = i.method === 'POST';
    const opName = isAssociate ? 'Associate' : 'Disassociate';
    const sources = i.multiRequests && i.multiRequests.length > 0
      ? i.multiRequests
      : [{ url, method: i.method, body: i.body ?? null }];
    return `// Xrm.WebApi.online.executeMultiple — ${opName} via CRUD descriptors.
// Each request describes one (target, relationship, related) tuple.

(async () => {
  const requests = [
${sources.map(_r => {
  // Best-effort parse from the existing URL/body. Real callers should
  // pass the parsed parts in directly.
  return `    {
      // target: { entityType: '<table>', id: '<guid>' },
      // relationship: '<schemaName>',
      // relatedEntity: { entityType: '<table>', id: '<guid>' },
      getMetadata: () => ({
        boundParameter: null,
        parameterTypes: {},
        operationType: 2,  // 2 = CRUD
        operationName: '${opName}',
      }),
    },`;
}).join('\n')}
  ];

  const results = await Xrm.WebApi.online.executeMultiple(requests);
  console.log(\`\${results.filter(r => r.ok).length}/\${results.length} succeeded\`);
})().catch(err => console.error('executeMultiple error:', err));
`;
  }

  // ── Retrieve Multiple (operationType: 2, operationName: 'RetrieveMultiple') ──
  if (isRetrieveMultiple) {
    return `// Xrm.WebApi.online.executeMultiple — batched retrieveMultiple.

(async () => {
  const requests = [
    {
      collection: '${i.built.entitySet || logical + 's'}',
      query: '${options}',
      getMetadata: () => ({
        boundParameter: null,
        parameterTypes: {},
        operationType: 2,  // 2 = CRUD
        operationName: 'RetrieveMultiple',
      }),
    },
  ];

  const results = await Xrm.WebApi.online.executeMultiple(requests);
  for (const r of results) {
    if (r.ok) {
      const data = await r.json();
      console.log(\`Returned \${data.value.length} row(s)\`);
      console.log(data.value);
    } else {
      console.error('Request failed:', r.status, r.statusText);
    }
  }
})().catch(err => console.error('executeMultiple error:', err));
`;
  }

  // ── Retrieve Single ──
  if (isRetrieveSingle) {
    return `// Xrm.WebApi.online.executeMultiple — batched retrieve of a single record.

(async () => {
  const requests = [
    {
      entity: { entityType: '${logical}', id: '${recordId ?? '<id>'}' },
      query: '${options}',
      getMetadata: () => ({
        boundParameter: null,
        parameterTypes: {},
        operationType: 2,  // 2 = CRUD
        operationName: 'Retrieve',
      }),
    },
  ];

  const results = await Xrm.WebApi.online.executeMultiple(requests);
  for (const r of results) {
    if (r.ok) {
      const record = await r.json();
      console.log('Got record:', record);
    } else {
      console.error('Retrieve failed:', r.status, r.statusText);
    }
  }
})().catch(err => console.error('executeMultiple error:', err));
`;
  }

  // ── Create / Update / Delete / Upsert ─────────────────────────
  if (isWrite) {
    const opName =
      i.method === 'POST'   ? 'Create' :
      i.method === 'PATCH'  ? 'Update' :
      i.method === 'DELETE' ? 'Delete' :
      'Update';
    const bodyLiteral = opName === 'Delete'
      ? ''
      : `\n      data: ${formatJsObject(i.body ?? {}, '      ')},\n`;
    // Honest fallback: if neither entitySet nor entityLogical is known we
    // emit a placeholder rather than guessing 'accounts'. The user can fill
    // in the right collection name (or the action detection above should
    // have caught this case — log a console warning if we get here).
    const collectionName =
      i.built.entitySet ||
      (logical ? `${logical}s` : '<entity-set>');
    const entityLine = recordId
      ? `      entity: { entityType: '${logical || '<entity-logical>'}', id: '${recordId}' },`
      : `      collection: '${collectionName}',`;
    return `// Xrm.WebApi.online.executeMultiple — batched ${opName.toLowerCase()} via CRUD descriptor.
//
// Extending: drop more descriptors into \`requests\` to do them in a batch.
// Pass \`false\` as the second arg to Xrm.WebApi.online.executeMultiple for
// Change Set (transactional) semantics; \`true\` for Faulted (continue on error).
//
// Notes:
//   • @odata.bind keys in data set lookup values; the value is the OData
//     entity-set URL fragment, e.g. '/contacts(<guid>)'.
//   • Multi-select choice columns are comma-separated integer strings.

(async () => {
  const requests = [
    {
${entityLine}${bodyLiteral}      getMetadata: () => ({
        boundParameter: null,
        parameterTypes: {},
        operationType: 2,  // 2 = CRUD
        operationName: '${opName}',
      }),
    },
  ];

  const results = await Xrm.WebApi.online.executeMultiple(requests);
  results.forEach((r, idx) => {
    if (r.ok) {
      console.log(\`Request \${idx} succeeded — \${r.status} \${r.statusText}\`);
    } else {
      console.error(\`Request \${idx} failed — \${r.status} \${r.statusText}\`);
    }
  });
})().catch(err => console.error('executeMultiple error:', err));
`;
  }
  return '';
}

// ────────────────────────────────────────────────────────────
// XMLHttpRequest — legacy / browsers without fetch
// ────────────────────────────────────────────────────────────
function genXhr(i: CodegenInputs): string {
  if (i.multiRequests && i.multiRequests.length > 1) return genFetchMulti(i);
  const fullUrlStr = fullUrl(i);
  const hdrs = stripAuth(i.headers);
  const setHeaders = Object.entries(hdrs)
    .filter(([, v]) => v)
    .map(([k, v]) => `req.setRequestHeader('${jsEscape(k)}', '${jsEscape(v)}');`)
    .join('\n');
  // Prefer relative URL when running inside Xrm (more portable across environments)
  return `// Dataverse Web API — XMLHttpRequest
// When running inside a model-driven app, prefer Xrm.Utility.getGlobalContext().getClientUrl()
// so the snippet works across environments without hardcoding the org URL.

(function () {
  var clientUrl = (typeof Xrm !== 'undefined' && Xrm.Utility && Xrm.Utility.getGlobalContext)
    ? Xrm.Utility.getGlobalContext().getClientUrl()
    : '${ENV.host ? 'https://' + ENV.host : ''}';
  var url = clientUrl + '${i.built.relativeUrl || (i.rawNextLink ?? '')}';

  var req = new XMLHttpRequest();
  req.open('${i.method}', url, true);
  req.setRequestHeader('Authorization', 'Bearer <access-token>');
${setHeaders ? '  ' + setHeaders.split('\n').join('\n  ') : ''}
  req.onreadystatechange = function () {
    if (req.readyState !== 4) return;
    req.onreadystatechange = null;
    if (req.status >= 200 && req.status < 300) {
      var data = req.status === 204 ? null : JSON.parse(req.responseText);
      console.log('Result:', data);
    } else {
      console.error('Dataverse ' + req.status + ':', req.responseText);
    }
  };
  req.send(${i.method === 'GET' || i.method === 'DELETE' ? '' : `JSON.stringify(${formatJsObject(i.body ?? {}, '  ')})`});
})();

// XMLHttpRequest example — full URL form (if you can't use Xrm.Utility):
//   var req = new XMLHttpRequest();
//   req.open('${i.method}', '${fullUrlStr}', true);
//   ...
`;
}

// ────────────────────────────────────────────────────────────
// Power Automate — STRUCTURED field spec (consumed by PowerAutomatePane)
// ────────────────────────────────────────────────────────────
//
// Returns a spec the UI can render as the field-by-field form that mirrors
// the Dataverse connector's "List rows" / "Get a row by ID" action — each
// row has a label, a read-only Input with the value, and a per-row Copy
// button. The text-only Monaco rendering (genPowerAutomate below) is kept
// as a fallback for environments that only show source.
export function generatePowerAutomateFields(i: CodegenInputs): PowerAutomateActionSpec {
  if (i.isNextLink && i.rawNextLink) {
    return {
      actionName: 'Follow @odata.nextLink',
      connector: 'Generic HTTP action',
      banner: 'The Dataverse "List rows" action won\'t accept an opaque nextLink — use the generic HTTP action instead. Most flows don\'t need this: set Maximum Pagination Size on "List rows" and the connector iterates pages automatically.',
      fields: [
        { label: 'Method', value: 'GET' },
        { label: 'URI', value: i.rawNextLink, multiline: true },
      ],
    };
  }

  const valueOf = (key: string): string => {
    const p = i.built.queryParts.find(x => x.key === key);
    return p ? decodeURIComponent(p.value) : '';
  };
  const select   = valueOf('$select');
  const filter   = valueOf('$filter');
  const orderby  = valueOf('$orderby');
  const expand   = valueOf('$expand');
  const top      = valueOf('$top');
  const apply    = valueOf('$apply');
  const countOn  = valueOf('$count') === 'true';
  const savedQ   = valueOf('savedQuery');
  const userQ    = valueOf('userQuery');

  let maxPageSize = '';
  const prefer = i.headers['Prefer'];
  if (prefer) {
    const m = prefer.match(/odata\.maxpagesize\s*=\s*(\d+)/i);
    if (m) maxPageSize = m[1];
  }

  const isRetrieveSingle = i.method === 'GET' && !!i.built.recordId;
  const isPredefined = !!savedQ || !!userQ;
  const isCreate = i.method === 'POST' && !i.built.recordId && !!i.body && !i.built.relativeUrl.includes('/Microsoft.Dynamics.CRM.') && !i.built.relativeUrl.endsWith('/$ref') && !i.built.relativeUrl.includes('/Merge');
  const isUpdate = i.method === 'PATCH' && !!i.built.recordId && !!i.body;
  const isDelete = i.method === 'DELETE' && !!i.built.recordId && !i.built.relativeUrl.endsWith('/$ref');
  const isAction = i.method === 'POST' && (i.built.relativeUrl.includes('/Microsoft.Dynamics.CRM.') || i.built.relativeUrl.endsWith('/Merge') || (!i.built.recordId && !i.built.entitySet && !!i.body));
  const isFunction = i.method === 'GET' && !i.built.recordId && !i.built.entitySet;

  // "Update a row" — Update / Upsert
  if (isUpdate) {
    const body = i.body ?? {};
    const fields: PowerAutomateField[] = [
      { label: 'Table name', value: i.built.entitySet, hint: 'Entity set name (plural — e.g. accounts)' },
      { label: 'Row ID',     value: i.built.recordId ?? '', hint: 'GUID of the record to update' },
    ];
    for (const [k, v] of Object.entries(body)) {
      if (k.endsWith('@odata.bind')) {
        const col = k.replace('@odata.bind', '');
        const m = typeof v === 'string' ? v.match(/\/([^/(]+)\((.*)\)/) : null;
        if (m) {
          const [, targetSet, guid] = m;
          fields.push({ label: col, value: guid, hint: `Lookup → ${targetSet}` });
        } else {
          fields.push({ label: col, value: String(v) });
        }
      } else {
        const val = typeof v === 'string' ? v : JSON.stringify(v);
        fields.push({ label: k, value: val });
      }
    }
    return {
      actionName: 'Update a row',
      connector: 'Microsoft Dataverse connector',
      banner: 'Open the "Update a row" action → pick the table, paste the Row ID, then fill only the columns you want to change. Leave others blank — they\'re left untouched.',
      fields,
      notes: [
        'PATCH semantics: only listed columns are updated. Empty fields stay as-is on the row.',
        'Lookup values: pick the related row from the searchable dropdown — the connector handles @odata.bind internally.',
        'Set a column to null via "Show advanced options" → null fields.',
      ],
    };
  }

  // "Delete a row" — Delete
  if (isDelete) {
    return {
      actionName: 'Delete a row',
      connector: 'Microsoft Dataverse connector',
      banner: 'Open the "Delete a row" action → pick the table + Row ID. There\'s no body and no "are you sure" — the row is gone on success.',
      fields: [
        { label: 'Table name', value: i.built.entitySet, hint: 'Entity set name (plural — e.g. accounts)' },
        { label: 'Row ID',     value: i.built.recordId ?? '', hint: 'GUID of the row to delete' },
      ],
      notes: [
        'Deletes cascade per relationship config — review cascade-delete rules first.',
        'For optimistic concurrency, send an If-Match etag header via the generic HTTP action; the Dataverse connector doesn\'t surface it.',
      ],
    };
  }

  // "Perform an unbound action" / "Perform a bound action" — Execute Action / Custom API / Custom Action
  if (isAction) {
    const isBound = !!i.built.recordId;
    const actionMatch = i.built.relativeUrl.match(/\/Microsoft\.Dynamics\.CRM\.([^/?(]+)|\/([^/?(]+)$/);
    const actionName = actionMatch ? (actionMatch[1] || actionMatch[2] || '') : '';
    const body = i.body ?? {};
    const paramFields: PowerAutomateField[] = Object.entries(body).map(([k, v]) => ({
      label: k,
      value: typeof v === 'string' ? v : JSON.stringify(v, null, 2),
      multiline: typeof v === 'object' && v !== null,
      hint: typeof v === 'object' && v !== null ? 'Complex / entity-typed param — paste this JSON into the parameter\'s value field' : undefined,
    }));
    return {
      actionName: isBound ? 'Perform a bound action' : 'Perform an unbound action',
      connector: 'Microsoft Dataverse connector',
      banner: isBound
        ? `Open the "Perform a bound action" → pick the table, paste the Row ID, then pick "${actionName}" from the action dropdown. Fill the parameters below.`
        : `Open the "Perform an unbound action" → pick "${actionName}" from the action dropdown (or paste this name). Fill the parameters below.`,
      fields: [
        { label: 'Action name', value: actionName, hint: 'Microsoft.Dynamics.CRM.<name> for OOB; <publisher>_<name> for custom' },
        ...(isBound
          ? [
              { label: 'Table name', value: i.built.entitySet, hint: 'Bound entity\'s entity set name' } as PowerAutomateField,
              { label: 'Row ID',     value: i.built.recordId ?? '', hint: 'GUID of the bound row' } as PowerAutomateField,
            ]
          : []),
        ...paramFields,
      ],
      notes: [
        'The Dataverse connector renders each parameter as a typed dynamic field — pick from dropdowns where the action accepts EntityReferences.',
        'Complex types (PrincipalAccess, OpportunityClose, etc.) are JSON objects — paste the full JSON into the parameter.',
      ],
    };
  }

  // Functions — no first-class connector action; recommend the HTTP action
  if (isFunction) {
    return {
      actionName: 'Invoke an HTTP request (function)',
      connector: 'Microsoft Dataverse connector',
      banner: 'The Dataverse connector doesn\'t expose OData functions directly. Use the connector\'s "Invoke an HTTP request" action (or the generic HTTP action) to GET the function URL.',
      fields: [
        { label: 'Method',   value: 'GET' },
        { label: 'Relative URI', value: i.built.relativeUrl.replace(/^\/api\/data\/v\d+\.\d+/, ''), hint: 'Strip the /api/data/v9.2 prefix — the connector adds it' },
        { label: 'Body',     value: '', hint: 'Functions have no body' },
      ],
      notes: [
        'Use parameter aliases (e.g. @p1) — they\'re recommended for any non-trivial value (long strings, GUIDs, DateTimeOffset).',
        'For bound functions, the URL is /<set>(<id>)/Microsoft.Dynamics.CRM.<name>(...). Inline params accepted but aliases preferred.',
      ],
      httpFallback: {
        method: 'GET',
        uri: `https://${ENV.host}${i.built.relativeUrl}`,
        note: 'Generic HTTP action — use this when "Invoke an HTTP request" isn\'t available in your connector version.',
      },
    };
  }

  // "Add a new row" — Create
  if (isCreate) {
    const body = i.body ?? {};
    // Split @odata.bind keys (lookups) from scalar values — the Dataverse
    // connector exposes each writable column as a separate dynamic field;
    // lookups go into the dropdown of the matching navigation property where
    // the user picks the related row, not pastes a /entitySet(guid) string.
    const fields: PowerAutomateField[] = [
      { label: 'Table name', value: i.built.entitySet, hint: 'Entity set name (plural — e.g. accounts)' },
    ];
    for (const [k, v] of Object.entries(body)) {
      if (k.endsWith('@odata.bind')) {
        const col = k.replace('@odata.bind', '');
        // Extract the GUID from "/<entitySet>(<guid>)"
        const m = typeof v === 'string' ? v.match(/\/([^/(]+)\((.*)\)/) : null;
        if (m) {
          const [, targetSet, guid] = m;
          fields.push({
            label: col,
            value: guid,
            hint: `Lookup → ${targetSet} · pick the row in the connector dropdown (or paste this GUID into the dynamic field)`,
          });
        } else {
          fields.push({ label: col, value: String(v) });
        }
      } else {
        const val = typeof v === 'string' ? v : JSON.stringify(v);
        fields.push({ label: k, value: val });
      }
    }
    return {
      actionName: 'Add a new row',
      connector: 'Microsoft Dataverse connector',
      banner: 'Open the "Add a new row" action in your flow → pick the table, then paste each value into the matching dynamic field. Lookup fields show the related table as a searchable dropdown — paste the GUID or pick the row from there.',
      fields,
      notes: [
        'The Dataverse connector exposes every writable column as a dynamic field — they appear only after you pick the table.',
        'Lookup values: pick the related row from the searchable dropdown. The connector handles the @odata.bind syntax internally.',
        'Multi-select choice: the action lets you select multiple options from a checkbox list.',
        'Choose "Show advanced options" to set MSCRM.SuppressDuplicateDetection or set the row owner.',
      ],
    };
  }

  // "Get a row by ID" — Retrieve Single
  if (isRetrieveSingle) {
    return {
      actionName: 'Get a row by ID',
      connector: 'Microsoft Dataverse connector',
      banner: 'Open the action in your flow → paste each value into the matching field. "Get a row by ID" only supports $select and $expand — Dataverse rejects $filter/$orderby/$top on a single-record URL.',
      fields: [
        { label: 'Table name',     value: i.built.entitySet, hint: 'Entity set name (plural — e.g. accounts)' },
        { label: 'Row ID',         value: i.built.recordId ?? '', hint: 'GUID of the record' },
        { label: 'Select Columns', value: select, hint: 'Logical names, comma-separated' },
        { label: 'Expand Query',   value: expand, multiline: true, hint: 'Nav property name with optional inner ($select=…) in parens' },
      ],
    };
  }

  // "List rows" with a Saved/User Query — Predefined Query mode
  if (isPredefined) {
    return {
      actionName: 'List rows  ·  Saved / User Query',
      connector: 'Microsoft Dataverse connector',
      banner: 'When you set Saved Query or User Query, the connector IGNORES the OData fields (Select / Filter / Sort / Expand) — those live inside the FetchXml the saved query embeds. Only Top Count and Maximum Pagination Size remain editable.',
      fields: [
        { label: 'Table name',              value: i.built.entitySet, hint: 'Entity set name (plural — e.g. accounts)' },
        ...(savedQ
          ? [{ label: 'Saved Query', value: savedQ, hint: 'GUID of a system saved query' } as PowerAutomateField]
          : [{ label: 'User Query',  value: userQ,  hint: 'GUID of a personal user query' } as PowerAutomateField]
        ),
        { label: 'Top Count',                value: top, hint: 'Optional — limits total rows returned' },
        { label: 'Maximum Pagination Size',  value: maxPageSize, hint: 'Per-page size; connector auto-iterates pages' },
      ],
      notes: [
        'Select Columns / Filter Rows / Sort By / Expand Query are intentionally NOT shown — the saved query owns them.',
        'Fetch Xml Query is mutually exclusive with Saved Query / User Query — leave it blank.',
      ],
    };
  }

  // "List rows" — the common case (Retrieve Multiple)
  const fields: PowerAutomateField[] = [
    { label: 'Table name',              value: i.built.entitySet, hint: 'Entity set name (plural — e.g. accounts)' },
    { label: 'Select Columns',          value: select,    hint: 'Logical names, comma-separated (no $select= prefix)' },
    { label: 'Filter Rows',             value: filter,    multiline: true, hint: 'OData filter expression (no $filter= prefix)' },
    { label: 'Sort By',                 value: orderby,   hint: '"col asc, col2 desc" (no $orderby= prefix)' },
    { label: 'Top Count',               value: top,       hint: 'Optional — total rows cap; leave blank to use pagination' },
    { label: 'Expand Query',            value: expand,    multiline: true, hint: 'Nav properties with optional inner ($select=…;$filter=…) in parens' },
    { label: 'Row Count',               value: countOn ? 'Yes' : 'No', hint: 'Adds @odata.count to the response when Yes' },
    { label: 'Maximum Pagination Size', value: maxPageSize, hint: 'Per-page size; connector auto-iterates pages' },
    { label: 'Fetch Xml Query',         value: '', hint: 'Alternative to OData params above — leave blank when using them' },
    { label: 'Saved Query',             value: '', hint: 'GUID — for system saved queries' },
    { label: 'User Query',              value: '', hint: 'GUID — for personal user queries' },
  ];
  const notes: string[] = [
    'Column names are LOGICAL names (e.g. "name", not "Account Name").',
    'Sort By takes "col asc, col2 desc" — no $orderby= prefix.',
    'Expand Query uses parens for inner options, e.g. primarycontactid($select=fullname,emailaddress1).',
    'Maximum Pagination Size enables auto-iteration of @odata.nextLink — leave Top Count blank when you want ALL matching rows.',
  ];
  if (apply) {
    notes.unshift('⚠ $apply (groupby / aggregate) is set — the "List rows" action doesn\'t have a field for this. Use the HTTP action below, or rewrite as FetchXml.');
  }

  return {
    actionName: 'List rows',
    connector: 'Microsoft Dataverse connector',
    banner: 'Open the "List rows" action in your flow and paste each value into the matching field. Empty fields can stay empty.',
    fields,
    notes,
    httpFallback: apply ? {
      method: i.method,
      uri: `https://${ENV.host}${i.built.relativeUrl}`,
      note: 'Use this generic HTTP action because $apply / aggregation isn\'t exposed by the "List rows" action. Replace <access-token> with a valid bearer token.',
    } : undefined,
  };
}

// ────────────────────────────────────────────────────────────
// Power Automate — Dataverse connector field-by-field cheatsheet
// ────────────────────────────────────────────────────────────
//
// Most Power Automate users don't want a JSON blob — they want the values
// they can paste straight into the Dataverse connector's "List rows" / "Get a
// row by ID" action fields. We emit a labeled key/value sheet in the order
// the connector's UI presents the fields, plus a fallback HTTP action shape
// for cases the high-level action can't handle (e.g. raw nextLink).
function genPowerAutomate(i: CodegenInputs): string {
  if (i.isNextLink && i.rawNextLink) {
    return `# Power Automate — follow an @odata.nextLink
#
# The Dataverse "List rows" action won't accept an opaque nextLink — it builds
# its own URL from the connector fields. To follow a nextLink directly, use
# the generic HTTP action:
#
#   Method:           GET
#   URI:              ${i.rawNextLink}
#   Authentication:   AAD (use a service-principal or OAuth connection)
#
# In most flows you don't need this — set "Maximum Pagination Size" in the
# Dataverse "List rows" action and the connector iterates pages automatically.
`;
  }

  // Pull each query option's value (decoded) from the built request.
  const valueOf = (key: string): string => {
    const p = i.built.queryParts.find(x => x.key === key);
    return p ? decodeURIComponent(p.value) : '';
  };
  const select   = valueOf('$select');
  const filter   = valueOf('$filter');
  const orderby  = valueOf('$orderby');
  const expand   = valueOf('$expand');
  const top      = valueOf('$top');
  const apply    = valueOf('$apply');
  const countOn  = valueOf('$count') === 'true';
  const savedQ   = valueOf('savedQuery');
  const userQ    = valueOf('userQuery');

  // Maximum Pagination Size from Prefer: odata.maxpagesize
  let maxPageSize = '';
  const prefer = i.headers['Prefer'];
  if (prefer) {
    const m = prefer.match(/odata\.maxpagesize\s*=\s*(\d+)/i);
    if (m) maxPageSize = m[1];
  }

  // Aligned "Field name  value" rendering — monospace in the editor lets users
  // double-click a value and Ctrl+C it straight into the flow designer.
  const row = (label: string, value: string): string =>
    `${label.padEnd(25)} ${value}`;
  const isRetrieveSingle = i.method === 'GET' && !!i.built.recordId;
  const isPredefined = !!savedQ || !!userQ;
  const isCreate = i.method === 'POST' && !i.built.recordId && !!i.body && !i.built.relativeUrl.includes('/Microsoft.Dynamics.CRM.') && !i.built.relativeUrl.endsWith('/$ref') && !i.built.relativeUrl.includes('/Merge');
  const isUpdate = i.method === 'PATCH' && !!i.built.recordId && !!i.body;
  const isDelete = i.method === 'DELETE' && !!i.built.recordId && !i.built.relativeUrl.endsWith('/$ref');
  const isAction = i.method === 'POST' && (i.built.relativeUrl.includes('/Microsoft.Dynamics.CRM.') || i.built.relativeUrl.endsWith('/Merge') || (!i.built.recordId && !i.built.entitySet && !!i.body));
  const isFunction = i.method === 'GET' && !i.built.recordId && !i.built.entitySet;

  // ── "Update a row" — Update / Upsert
  if (isUpdate) {
    const body = i.body ?? {};
    const lines: string[] = [];
    lines.push(row('Table name', i.built.entitySet));
    lines.push(row('Row ID', i.built.recordId ?? ''));
    for (const [k, v] of Object.entries(body)) {
      const val = typeof v === 'string' ? v : JSON.stringify(v);
      lines.push(row(k, val));
    }
    return `# Power Automate — Dataverse connector  ⟶  "Update a row" action
# Paste each value into the matching dynamic field. Leave non-listed columns alone — PATCH only touches what's in the body.

${lines.join('\n')}

# ── Tips ──
#  • Lookup columns: pick from the searchable dropdown — the connector emits @odata.bind for you.
#  • Set a column to null via "Show advanced options" → null fields list.
#  • For optimistic concurrency, fall back to the HTTP action with an If-Match etag header.
`;
  }

  // ── "Delete a row" — Delete
  if (isDelete) {
    return `# Power Automate — Dataverse connector  ⟶  "Delete a row" action
# Pick the table + paste the Row ID. No body, no "are you sure".

${row('Table name', i.built.entitySet)}
${row('Row ID', i.built.recordId ?? '')}

# ── Tips ──
#  • Deletes cascade per relationship config — review cascade-delete rules first.
#  • Etag concurrency requires the HTTP action; the connector doesn't surface If-Match.
`;
  }

  // ── "Perform an (un)bound action" — Action / Custom API / Merge / Custom Action
  if (isAction) {
    const isBound = !!i.built.recordId;
    const actionMatch = i.built.relativeUrl.match(/\/Microsoft\.Dynamics\.CRM\.([^/?(]+)|\/([^/?(]+)$/);
    const actionName = actionMatch ? (actionMatch[1] || actionMatch[2] || '') : '';
    const body = i.body ?? {};
    const paramLines = Object.entries(body).map(([k, v]) => {
      const val = typeof v === 'string' ? v : JSON.stringify(v);
      return row(k, val.length > 60 ? val.slice(0, 57) + '…' : val);
    });
    const actionLabel = isBound ? '"Perform a bound action"' : '"Perform an unbound action"';
    return `# Power Automate — Dataverse connector  ⟶  ${actionLabel}
# Pick the action from the dropdown — the connector renders typed parameter fields per metadata.

${row('Action name', actionName)}
${isBound ? row('Table name', i.built.entitySet) + '\n' + row('Row ID', i.built.recordId ?? '') + '\n' : ''}${paramLines.join('\n')}

# ── Tips ──
#  • Complex / entity-typed parameters render as nested JSON fields — paste the full object.
#  • For OOB actions use the Microsoft.Dynamics.CRM.<name> form; for custom APIs / actions use the <publisher>_<name> form.
`;
  }

  // ── Functions — fall back to HTTP action
  if (isFunction) {
    return `# Power Automate — function call
# The Dataverse connector doesn't expose OData functions directly. Use the
# "Invoke an HTTP request" action (Dataverse connector) or the generic HTTP action.

  Method:    GET
  URI:       ${i.built.relativeUrl}
  Headers:   { "Authorization": "Bearer <access-token>" }

# ── Tips ──
#  • Use parameter aliases (@p1) — recommended for long values, GUIDs, DateTimeOffset.
#  • Bound functions: /<set>(<id>)/Microsoft.Dynamics.CRM.<name>(...).
`;
  }

  // ── "Add a new row" — Create
  if (isCreate) {
    const body = i.body ?? {};
    const lines: string[] = [];
    lines.push(row('Table name', i.built.entitySet));
    for (const [k, v] of Object.entries(body)) {
      if (k.endsWith('@odata.bind')) {
        const col = k.replace('@odata.bind', '');
        const m = typeof v === 'string' ? v.match(/\/([^/(]+)\((.*)\)/) : null;
        if (m) {
          const [, targetSet, guid] = m;
          lines.push(row(col, `${guid}   (→ ${targetSet})`));
        } else {
          lines.push(row(col, String(v)));
        }
      } else {
        lines.push(row(k, typeof v === 'string' ? v : JSON.stringify(v)));
      }
    }
    return `# Power Automate — Dataverse connector  ⟶  "Add a new row" action
# Paste each value into the matching dynamic field. The Dataverse connector
# discovers writable columns from the table you pick and renders them inline.

${lines.join('\n')}

# ── Tips ──
#  • Lookup dynamic fields show the related table as a searchable dropdown —
#    the GUIDs above can be pasted directly or used to pick the row.
#  • Multi-select choice: the connector exposes a multi-checkbox UI; values
#    above are comma-separated integers.
#  • Show advanced options to set Caller ID for impersonation or toggle
#    duplicate detection (MSCRM.SuppressDuplicateDetection).
`;
  }

  // ── "Get a row by ID" — Retrieve Single
  if (isRetrieveSingle) {
    return `# Power Automate — Dataverse connector  ⟶  "Get a row by ID" action
# Paste each value into the matching field. Leave blank rows alone.
# (Open the action in your flow → switch each parameter to the value below.)

${row('Table name',      i.built.entitySet)}
${row('Row ID',          i.built.recordId ?? '')}
${row('Select Columns',  select)}
${row('Expand Query',    expand)}

# ── Tip ──
#  • "Get a row by ID" only supports $select and $expand — those are the only
#    OData options Dataverse allows when the URL has the (id) segment.
#  • For impersonation, set the "Caller ID" advanced parameter to the user's GUID.
`;
  }

  // ── "List rows" with a Saved or User query (Predefined Query mode)
  if (isPredefined) {
    return `# Power Automate — Dataverse connector  ⟶  "List rows" action (with a Saved/User query)
# Paste each value into the matching field.
# When you set Saved Query / User Query, the Dataverse connector IGNORES the OData
# fields (Select / Filter / Sort / Expand) — those live inside the FetchXml the
# saved query embeds. Top Count and Pagination still apply.

${row('Table name',              i.built.entitySet)}
${savedQ ? row('Saved Query',    savedQ) : row('User Query', userQ)}
${row('Top Count',               top)}
${row('Maximum Pagination Size', maxPageSize)}

# Leave these blank — the saved query owns them:
${row('Select Columns',          '')}
${row('Filter Rows',             '')}
${row('Sort By',                 '')}
${row('Expand Query',            '')}
${row('Row Count',               'No')}
`;
  }

  // ── "List rows" — Retrieve Multiple (the common case)
  const apply_warning = apply ? `\n# ⚠ $apply (groupby/aggregate) — the "List rows" action doesn't expose a field for this.
#   Use the generic HTTP action below to send the request directly, or move to FetchXml.\n` : '';
  return `# Power Automate — Dataverse connector  ⟶  "List rows" action
# Paste each value into the matching field. Leave blank rows alone.
# (Open the action in your flow → switch each parameter to the value below.)

${row('Table name',              i.built.entitySet)}
${row('Select Columns',          select)}
${row('Filter Rows',             filter)}
${row('Sort By',                 orderby)}
${row('Top Count',               top)}
${row('Expand Query',            expand)}
${row('Row Count',               countOn ? 'Yes' : 'No')}
${row('Maximum Pagination Size', maxPageSize)}
${row('Fetch Xml Query',         '')}
${row('Saved Query',             '')}
${row('User Query',              '')}
${apply_warning}
# ── Notes ──
#  • Column names are LOGICAL names (e.g. "name" not "Account Name").
#  • Filter Rows takes the OData expression — no "$filter=" prefix.
#  • Sort By takes "col asc, col2 desc" — no "$orderby=" prefix.
#  • Expand Query takes the inner clauses in parens, e.g.
#      primarycontactid($select=fullname,emailaddress1)
#  • Maximum Pagination Size enables the connector to follow nextLinks automatically;
#    leave Top Count blank when using pagination to fetch ALL matching rows.

# ── Fallback — generic HTTP action ──
# Use this only when the "List rows" action can't express the request
# (e.g. $apply, raw nextLink, custom URL params). Replace <access-token> with
# a token from your OAuth / Service Principal step.

  Method:   ${i.method}
  URI:      https://${ENV.host}${i.built.relativeUrl}
  Headers:  { "Authorization": "Bearer <access-token>" }
`;
}

// ────────────────────────────────────────────────────────────
// C# — HttpClient
// ────────────────────────────────────────────────────────────
function genCsharp(i: CodegenInputs): string {
  if (i.multiRequests && i.multiRequests.length > 1) {
    const reqs = i.multiRequests;
    const headerLines = Object.entries(i.headers)
      .filter(([, v]) => v)
      .filter(([k]) => k.toLowerCase() !== 'authorization')
      .map(([k, v]) => `http.DefaultRequestHeaders.TryAddWithoutValidation("${k}", "${v.replace(/"/g, '\\"')}");`)
      .join('\n');
    const requestObjects = reqs.map(r => {
      const url = `https://${ENV.host}${r.relativeUrl}`;
      const desc = r.description ? `        // ${r.description}\n` : '';
      const body = r.body
        ? `, content: @"${csharpVerbatimEscape(JSON.stringify(r.body))}"`
        : '';
      return `${desc}        new { method = HttpMethod.${methodCsharp(r.method)}, url = "${url}"${body} },`;
    }).join('\n');
    return `// .NET 8 — HttpClient (multi-request)
// Associate / Disassociate over a collection fires one call per related id.

using System.Net.Http;
using System.Net.Http.Headers;

var http = new HttpClient { BaseAddress = new Uri("https://${ENV.host}/") };
http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", "<access-token>");
${headerLines}

var requests = new[] {
${requestObjects}
};

int ok = 0;
foreach (var r in requests)
{
    var req = new HttpRequestMessage((HttpMethod)((dynamic)r).method, (string)((dynamic)r).url);
    var content = ((dynamic)r).content as string;
    if (!string.IsNullOrEmpty(content))
        req.Content = new StringContent(content, System.Text.Encoding.UTF8, "application/json");
    using var res = await http.SendAsync(req);
    if (res.IsSuccessStatusCode) ok++;
    else Console.WriteLine($"{r} → {(int)res.StatusCode} {res.ReasonPhrase}");
}
Console.WriteLine($"{ok}/{requests.Length} requests succeeded");
`;
  }
  const url = fullUrl(i);
  const headerLines = Object.entries(i.headers)
    .filter(([, v]) => v)
    .filter(([k]) => k.toLowerCase() !== 'authorization')
    .map(([k, v]) => `http.DefaultRequestHeaders.TryAddWithoutValidation("${k}", "${v.replace(/"/g, '\\"')}");`)
    .join('\n');
  return `// .NET 8 — HttpClient. For production use Microsoft.PowerPlatform.Dataverse.Client
// (ServiceClient) instead — it handles auth, retries, and OData parsing for you.

using System.Net.Http;
using System.Net.Http.Headers;

var http = new HttpClient { BaseAddress = new Uri("https://${ENV.host}/") };
http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", "<access-token>");
${headerLines}

var req = new HttpRequestMessage(HttpMethod.${methodCsharp(i.method)}, "${url}");
${i.method !== 'GET' ? `req.Content = new StringContent(@"${csharpVerbatimEscape(JSON.stringify(i.body ?? {}, null, 2))}", System.Text.Encoding.UTF8, "application/json");\n` : ''}
using var res = await http.SendAsync(req);
res.EnsureSuccessStatusCode();
var json = await res.Content.ReadAsStringAsync();
Console.WriteLine(json);
`;
}

/** Escape a JSON string for inclusion inside a C# verbatim string literal (@"…"). */
function csharpVerbatimEscape(s: string): string {
  // In an @"…" verbatim literal, the only character that needs escaping is `"` → `""`.
  return s.replace(/"/g, '""');
}

function methodCsharp(m: string): string {
  switch (m.toUpperCase()) {
    case 'GET': return 'Get';
    case 'POST': return 'Post';
    case 'PATCH': return 'Patch';
    case 'PUT': return 'Put';
    case 'DELETE': return 'Delete';
    default: return 'Get';
  }
}

// ────────────────────────────────────────────────────────────
// PowerShell — Invoke-RestMethod
// ────────────────────────────────────────────────────────────
function genPowershell(i: CodegenInputs): string {
  if (i.multiRequests && i.multiRequests.length > 1) {
    const reqs = i.multiRequests;
    const headerLines = Object.entries(i.headers)
      .filter(([, v]) => v)
      .filter(([k]) => k.toLowerCase() !== 'authorization')
      .map(([k, v]) => `  "${k}" = "${v.replace(/"/g, '`"')}"`)
      .join('\n');
    const lines = reqs.map(r => {
      const url = `https://${ENV.host}${r.relativeUrl}`;
      const body = r.body ? ` -ContentType 'application/json' -Body '${JSON.stringify(r.body).replace(/'/g, "''")}'` : '';
      return `Invoke-RestMethod -Method ${r.method} -Uri "${url}" -Headers $headers${body}`;
    }).join('\n');
    return `# PowerShell — Invoke-RestMethod (multi-request)
$token = "<access-token>"
$headers = @{
  Authorization = "Bearer $token"
${headerLines}
}

${lines}
`;
  }
  const url = fullUrl(i);
  const headerLines = Object.entries(i.headers)
    .filter(([, v]) => v)
    .filter(([k]) => k.toLowerCase() !== 'authorization')
    .map(([k, v]) => `  "${k}" = "${v.replace(/"/g, '`"')}"`)
    .join('\n');
  return `# PowerShell — Invoke-RestMethod
# Replace <access-token> with a real bearer token (Get-AzAccessToken / msal-token).

$token = "<access-token>"
$headers = @{
  Authorization = "Bearer $token"
${headerLines}
}

${i.method !== 'GET' ? `$body = @'\n${JSON.stringify(i.body ?? {}, null, 2)}\n'@\n\n` : ''}$response = Invoke-RestMethod \`
  -Method ${i.method} \`
  -Uri "${url}" \`
  -Headers $headers${i.method !== 'GET' ? ' `\n  -ContentType "application/json" `\n  -Body $body' : ''}

$response | ConvertTo-Json -Depth 8
`;
}

// ────────────────────────────────────────────────────────────
// cURL — shell / CI scripts
// ────────────────────────────────────────────────────────────
function genCurl(i: CodegenInputs): string {
  if (i.multiRequests && i.multiRequests.length > 1) {
    const reqs = i.multiRequests;
    const hdrs = stripAuth(i.headers);
    const hdrLines = Object.entries(hdrs)
      .filter(([, v]) => v)
      .map(([k, v]) => `  -H "${k}: ${v.replace(/"/g, '\\"')}"`)
      .join(' \\\n');
    const lines = reqs.map(r => {
      const url = `https://${ENV.host}${r.relativeUrl}`;
      const desc = r.description ? `# ${r.description}\n` : '';
      const body = r.body ? ` \\\n  --data-raw '${JSON.stringify(r.body).replace(/'/g, "'\\''")}'` : '';
      return `${desc}curl -X ${r.method} --globoff "${url}" \\\n  -H "Authorization: Bearer <access-token>" \\\n${hdrLines}${body}`;
    }).join('\n\n');
    return `# cURL — multi-request (Associate / Disassociate over a collection)\n# Each association is a separate request per the docs.\n\n${lines}\n`;
  }
  const url = fullUrl(i);
  const hdrs = stripAuth(i.headers);
  const hdrLines = Object.entries(hdrs)
    .filter(([, v]) => v)
    .map(([k, v]) => `  -H "${k}: ${v.replace(/"/g, '\\"')}"`)
    .join(' \\\n');
  const body = i.method !== 'GET' && i.method !== 'DELETE'
    ? ` \\\n  --data-raw '${JSON.stringify(i.body ?? {}).replace(/'/g, "'\\''")}'`
    : '';
  return `# cURL — shell context
# --globoff prevents curl from interpreting the literal $ in OData query options
# --compressed asks the server for gzipped responses

curl -X ${i.method} \\
  --globoff \\
  --compressed \\
  "${url}" \\
  -H "Authorization: Bearer <access-token>" \\
${hdrLines}${body}
`;
}

// ────────────────────────────────────────────────────────────
// JSON — raw description (for Postman / ad-hoc tooling)
// ────────────────────────────────────────────────────────────
function genJson(i: CodegenInputs): string {
  if (i.multiRequests && i.multiRequests.length > 1) {
    return JSON.stringify({
      mode: 'multi-request',
      headers: i.headers,
      requests: i.multiRequests.map(r => ({
        method: r.method,
        url: `https://${ENV.host}${r.relativeUrl}`,
        relativeUrl: r.relativeUrl,
        body: r.body,
        description: r.description,
      })),
    }, null, 2);
  }
  const out: Record<string, unknown> = {
    method: i.method,
    url: fullUrl(i),
    relativeUrl: i.built.relativeUrl,
    entityLogical: i.built.entityLogical,
    entitySet: i.built.entitySet,
    headers: i.headers,
  };
  if (i.method !== 'GET' && i.body) out.body = i.body;
  return JSON.stringify(out, null, 2);
}
