// Turn a parsed curl command into a request for the matching mode (#271).
//
// The pasted text is untrusted: the URL must be a Dataverse Web API URL, the
// entity set must be a table in this environment, record keys must be GUIDs,
// and credential headers are dropped. Nothing is sent from here — the mode
// the request lands in shows it, and Execute is still the user's call.

import {
  parseODataUrl,
  type EntitySetResolver,
  type LoadTable,
  type ParsedRequest,
} from '../odataParser';
import { parseRawBody } from '../fieldBody';
import { impersonationVar, isCredentialHeader } from '../sensitiveHeaders';
import { decodeParamFromBody } from '../executeBuilders';
import { emptyPrefer, type PreferSpec } from '../../editors/PreferEditor';
import {
  defaultBypassOptions,
  type BypassOptions,
  type ConcurrencyMode,
  type DeleteScope,
} from '../../state/writeState';
import type { CsdlAction } from '../../metadata/actionsCsdl';
import type { ExpandSpec } from '../../editors/ExpandEditor';
import { newGroup, newItem, type BatchGroup } from '../batch/batchModel';
import { parseBatchRequest } from '../batch/multipart';
import { parseCurl, type CurlHeader } from './curlCommand';

/** A header row to merge into the mode's header list. */
export type ImportedHeader = CurlHeader;

interface WriteBits {
  table: string;
  /** Pretty-printed JSON object. */
  rawBody: string;
  prefer: PreferSpec;
  returnSelect: string[];
  bypass: BypassOptions;
  headers: ImportedHeader[];
}

/** What lands in the mode, keyed by the registry id of the mode it opens. */
export type ImportedRequest =
  | {
      modeId: 'retrieve-multiple';
      parsed: ParsedRequest;
      prefer: PreferSpec;
      headers: ImportedHeader[];
    }
  | {
      modeId: 'retrieve-single';
      table: string;
      recordId: string;
      select: string[];
      expand: ExpandSpec[];
      prefer: PreferSpec;
      headers: ImportedHeader[];
    }
  | ({ modeId: 'create' } & WriteBits)
  | ({ modeId: 'update'; recordId: string; concurrency: ConcurrencyMode } & WriteBits)
  | ({ modeId: 'upsert'; recordId: string; concurrency: ConcurrencyMode } & WriteBits)
  | {
      modeId: 'delete';
      table: string;
      recordId: string;
      scope: DeleteScope;
      concurrency: ConcurrencyMode;
      bypass: BypassOptions;
      headers: ImportedHeader[];
    }
  | {
      modeId: 'exec-action' | 'exec-customapi' | 'exec-customaction';
      actionName: string;
      boundRecordId: string | null;
      paramValues: Record<string, unknown>;
      headers: ImportedHeader[];
    }
  | { modeId: 'batch'; groups: BatchGroup[]; continueOnError: boolean };

export type ImportModeId = ImportedRequest['modeId'];

export interface CurlImportContext {
  /** Host of the connected environment (`contoso.crm.dynamics.com`), '' standalone. */
  host: string;
  resolveEntitySet: EntitySetResolver;
  loadTable?: LoadTable;
  findAction?: (name: string) => Promise<CsdlAction | undefined>;
}

export interface CurlImportResult {
  request?: ImportedRequest;
  method?: string;
  /** Path after the API base, for the summary (`accounts(…)`). */
  path?: string;
  /** What was left out or changed on the way in. */
  notes: string[];
  warnings: string[];
  errors: string[];
}

const GUID_RE = /^\{?[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\}?$/i;
const IDENT_RE = /^[A-Za-z_][A-Za-z0-9_]*$/;
const API_RE = /^\/api\/data\/v9\.\d+\/(.*)$/i;
const NS = 'Microsoft.Dynamics.CRM.';

/** Headers a browser adds that say nothing about the Dataverse request. */
const BROWSER_HEADERS = new Set([
  'user-agent',
  'referer',
  'origin',
  'accept-language',
  'accept-encoding',
  'priority',
  'dnt',
  'connection',
  'host',
  'content-length',
  'cache-control',
  'pragma',
  'x-requested-with',
  'clienthost',
  'x-ms-user-agent',
  'x-ms-client-request-id',
  'x-ms-client-session-id',
  'x-ms-app-name',
  'x-ms-correlation-id',
]);

const ANNOTATION_FLAGS: Record<string, keyof PreferSpec> = {
  'odata.community.display.v1.formattedvalue': 'formattedValues',
  'microsoft.dynamics.crm.lookuplogicalname': 'lookupLogicalNames',
  'microsoft.dynamics.crm.associatednavigationproperty': 'associatedNavProperty',
  'microsoft.dynamics.crm.totalrecordcount': 'totalRecordCount',
  'microsoft.dynamics.crm.totalrecordcountlimitexceeded': 'totalRecordCount',
};

/** Split a header value on commas that aren't inside double quotes. */
function splitList(value: string): string[] {
  const out: string[] = [];
  let cur = '';
  let quoted = false;
  for (const c of value) {
    if (c === '"') quoted = !quoted;
    if (c === ',' && !quoted) {
      out.push(cur.trim());
      cur = '';
    } else cur += c;
  }
  if (cur.trim()) out.push(cur.trim());
  return out.filter(Boolean);
}

/** Read a Prefer header back into the Prefer pane's switches. */
export function parsePrefer(value: string | null): {
  spec: PreferSpec;
  continueOnError: boolean;
  unknown: string[];
} {
  const spec: PreferSpec = {
    ...emptyPrefer(),
    formattedValues: false,
  };
  const unknown: string[] = [];
  let continueOnError = false;
  if (!value) return { spec, continueOnError, unknown };
  for (const pref of splitList(value)) {
    const lower = pref.toLowerCase();
    const anno = /^odata\.include-annotations\s*=\s*"?([^"]*)"?$/i.exec(pref);
    if (anno) {
      for (const a of anno[1].split(',').map((s) => s.trim())) {
        if (a === '*') spec.allAnnotations = true;
        else if (ANNOTATION_FLAGS[a.toLowerCase()])
          (spec[ANNOTATION_FLAGS[a.toLowerCase()]] as boolean) = true;
        else if (a) unknown.push(`odata.include-annotations="${a}"`);
      }
    } else if (lower === 'return=representation') spec.returnRepresentation = true;
    else if (lower === 'return=minimal') spec.returnRepresentation = false;
    else if (/^odata\.maxpagesize\s*=\s*\d+$/.test(lower))
      spec.maxpagesize = Number(lower.split('=')[1]);
    else if (lower === 'odata.continue-on-error') continueOnError = true;
    else unknown.push(pref);
  }
  return { spec, continueOnError, unknown };
}

const BYPASS_LOGIC = 'mscrm.bypassbusinesslogicexecution';
const BYPASS_STEPS = 'mscrm.bypassbusinesslogicexecutionstepids';
const BYPASS_LEGACY = 'mscrm.bypasscustompluginexecution';
const SUPPRESS_FLOWS = 'mscrm.suppresscallbackregistrationexpanderjob';

/**
 * Read the bypass headers back into the Bypass pane. The pane holds one
 * business-logic choice (sync / async / both, step ids, or the legacy
 * header), so a combination it can't show stays as header rows, as sent.
 */
function readBypass(headers: CurlHeader[]): {
  bypass: BypassOptions;
  rest: CurlHeader[];
  kept: boolean;
} {
  const bypass = defaultBypassOptions();
  const isLogic = (h: CurlHeader) =>
    [BYPASS_LOGIC, BYPASS_STEPS, BYPASS_LEGACY].includes(h.name.toLowerCase());
  const logic = headers.filter(isLogic);
  const flows = headers.filter((h) => h.name.toLowerCase() === SUPPRESS_FLOWS);
  let rest = headers.filter((h) => !isLogic(h) && h.name.toLowerCase() !== SUPPRESS_FLOWS);
  let readable = logic.length <= 1 && flows.length <= 1;
  if (readable && logic.length === 1) {
    const h = logic[0];
    const name = h.name.toLowerCase();
    const v = h.value.replace(/\s+/g, '').toLowerCase();
    const parts = new Set(v.split(','));
    if (
      name === BYPASS_LOGIC &&
      parts.size === 2 &&
      parts.has('customsync') &&
      parts.has('customasync')
    )
      bypass.businessLogic = 'both';
    else if (name === BYPASS_LOGIC && v === 'customsync') bypass.businessLogic = 'sync';
    else if (name === BYPASS_LOGIC && v === 'customasync') bypass.businessLogic = 'async';
    else if (name === BYPASS_LEGACY && v === 'true') {
      bypass.businessLogic = 'sync';
      bypass.useLegacyHeader = true;
    } else if (name === BYPASS_STEPS) {
      const ids = h.value.split(',').map((s) => s.trim());
      if (ids.every((id) => GUID_RE.test(id))) {
        bypass.businessLogic = 'steps';
        bypass.stepIds = ids.map(cleanGuid);
      } else readable = false;
    } else readable = false;
  }
  if (readable && flows.length === 1) {
    if (flows[0].value.trim().toLowerCase() === 'true') bypass.suppressFlows = true;
    else readable = false;
  }
  if (!readable) {
    rest = headers;
    return { bypass: defaultBypassOptions(), rest, kept: logic.length + flows.length > 0 };
  }
  return { bypass, rest, kept: false };
}

function concurrencyOf(ifMatch: string | null, ifNoneMatch: string | null): ConcurrencyMode {
  if (ifNoneMatch?.trim() === '*') return { kind: 'create-only' };
  if (ifMatch?.trim() === '*') return { kind: 'update-only' };
  if (ifMatch) return { kind: 'etag', etag: ifMatch.trim() };
  return { kind: 'none' };
}

const cleanGuid = (s: string) => s.replace(/^\{|\}$/g, '').toLowerCase();

function safeDecode(s: string): string {
  try {
    return decodeURIComponent(s);
  } catch {
    return s;
  }
}

/** `select=a,b` → ['a','b'], from a raw query string. */
function queryParams(query: string): Map<string, string> {
  const out = new Map<string, string>();
  for (const pair of query.split('&')) {
    if (!pair) continue;
    const eq = pair.indexOf('=');
    const k = safeDecode(eq < 0 ? pair : pair.slice(0, eq));
    out.set(k, safeDecode(eq < 0 ? '' : pair.slice(eq + 1)));
  }
  return out;
}

/** Read a pasted curl command into the request it opens. */
export async function importCurl(text: string, ctx: CurlImportContext): Promise<CurlImportResult> {
  const notes: string[] = [];
  const warnings: string[] = [];
  const errors: string[] = [];
  const fail = (message: string): CurlImportResult => ({
    notes,
    warnings,
    errors: [...errors, message],
  });

  const p = parseCurl(text);
  if (!p.ok) return fail(p.error);
  const cmd = p.command;
  notes.push(...cmd.notes);
  const method = cmd.method;
  if (!['GET', 'POST', 'PATCH', 'PUT', 'DELETE'].includes(method))
    return fail(`${method} isn't a Dataverse Web API method.`);

  // ── URL: a Dataverse Web API URL, absolute (https) or from the API base ──
  const abs = /^(https?):\/\/([^/?#]+)(.*)$/i.exec(cmd.url);
  if (!abs && !cmd.url.startsWith('/'))
    return fail(`"${cmd.url}" isn't a Dataverse Web API URL (https://<org>/api/data/v9.x/…).`);
  if (abs && abs[1].toLowerCase() !== 'https')
    return fail('The URL must use https: Dataverse only answers over https.');
  const host = abs ? abs[2].toLowerCase() : '';
  const pathAndQuery = (abs ? abs[3] : cmd.url).replace(/#.*$/, '');
  const qAt = pathAndQuery.indexOf('?');
  const fullPath = qAt < 0 ? pathAndQuery : pathAndQuery.slice(0, qAt);
  const query = qAt < 0 ? '' : pathAndQuery.slice(qAt + 1);
  const api = API_RE.exec(fullPath);
  if (!api) return fail(`The URL path isn't under /api/data/v9.x/ — "${fullPath || '/'}".`);
  const path = api[1].replace(/\/+$/, '');
  if (host && ctx.host && host !== ctx.host.toLowerCase()) {
    warnings.push(
      `The command is for ${host}; it opens against the environment you're connected to (${ctx.host}).`,
    );
  }
  if (path.split('/').some((s) => s === '..' || s === '.'))
    return fail('The URL path has "." or ".." segments.');

  // ── Headers ──
  const kept: CurlHeader[] = [];
  const browser: string[] = [];
  let prefer: string | null = null;
  let ifMatch: string | null = null;
  let ifNoneMatch: string | null = null;
  let contentType: string | null = null;
  for (const h of cmd.headers) {
    const name = h.name.toLowerCase();
    if (!/^[!#$%&'*+.^_`|~0-9A-Za-z-]+$/.test(h.name))
      return fail(`"${h.name}" isn't a header name.`);
    // eslint-disable-next-line no-control-regex
    if (/[\x00-\x1f\x7f]/.test(h.value))
      return fail(`The ${h.name} header has a control character.`);
    if (isCredentialHeader(h.name)) notes.push(`Left out ${h.name}: ToolBox signs you in.`);
    else if (BROWSER_HEADERS.has(name) || name.startsWith('sec-')) browser.push(h.name);
    else if (name === 'prefer') prefer = prefer ? `${prefer}, ${h.value}` : h.value;
    else if (name === 'if-match') ifMatch = h.value;
    else if (name === 'if-none-match') ifNoneMatch = h.value;
    else {
      if (name === 'content-type') contentType = h.value;
      kept.push(h);
    }
  }
  if (browser.length) notes.push(`Left out browser headers: ${browser.join(', ')}.`);
  const pref = parsePrefer(prefer);
  // Bypass headers fill the write modes' Bypass pane; reads and actions keep them as rows.
  const { bypass, rest, kept: bypassAsRows } = readBypass(kept);
  const bypassNote = () => {
    if (bypassAsRows)
      warnings.push(
        "Kept the bypass headers as header rows: the Bypass pane can't show this combination.",
      );
  };
  // Headers a mode has no place for, said out loud instead of dropped quietly.
  const leftOut = (what: string, why: string) => warnings.push(`Left out ${what}: ${why}.`);
  const noPreconditions = (why: string) => {
    if (ifMatch) leftOut(`If-Match: ${ifMatch}`, why);
    if (ifNoneMatch) leftOut(`If-None-Match: ${ifNoneMatch}`, why);
  };
  const noPageSize = () => {
    if (pref.spec.maxpagesize != null)
      leftOut(`Prefer: odata.maxpagesize=${pref.spec.maxpagesize}`, 'only queries page');
  };
  // Reads keep If-None-Match (304 Not Modified on a row) as a header row.
  const readRows = () => {
    if (ifMatch) leftOut(`If-Match: ${ifMatch}`, 'reads use If-None-Match');
    return ifNoneMatch ? [...kept, { name: 'If-None-Match', value: ifNoneMatch }] : kept;
  };

  const summary = { method, path: path + (query ? `?${query}` : '') };
  const done = (request: ImportedRequest): CurlImportResult => ({
    request,
    ...summary,
    notes,
    warnings,
    errors,
  });
  const unsupported = (what: string, instead: string) =>
    fail(`Import doesn't open ${what} yet. ${instead}`);
  const prefNotes = (allowed: 'read' | 'write' | 'none') => {
    if (pref.unknown.length) warnings.push(`Left out Prefer: ${pref.unknown.join(', ')}.`);
    if (pref.continueOnError)
      warnings.push('Left out Prefer: odata.continue-on-error (only $batch uses it).');
    if (allowed === 'none' && prefer)
      warnings.push(`Left out Prefer: ${prefer} (this mode doesn't send one).`);
    if (allowed === 'read' && pref.spec.returnRepresentation)
      warnings.push('Left out Prefer: return=representation (reads return the rows anyway).');
  };
  const noBody = () => {
    if (cmd.body !== undefined && cmd.body !== '')
      warnings.push(`Left out the body: a ${method} request doesn't send one.`);
  };
  const jsonBody = (required: boolean): string | null => {
    if (cmd.body === undefined || cmd.body === '') {
      if (required) {
        errors.push(`A ${method} request needs a JSON body.`);
        return null;
      }
      return '{}';
    }
    if (contentType && !/json/i.test(contentType)) {
      errors.push(`The body is ${contentType}, not JSON.`);
      return null;
    }
    const r = parseRawBody(cmd.body);
    if (!r.ok) {
      errors.push(`The body isn't a JSON object: ${r.error}`);
      return null;
    }
    return JSON.stringify(r.body, null, 2);
  };
  const writeQuery = () => {
    const params = queryParams(query);
    const select = params.get('$select');
    params.delete('$select');
    if (params.size) warnings.push(`Left out ${[...params.keys()].join(', ')} from the URL.`);
    // A lookup is selected on the wire as `_<column>_value`; the Return
    // columns picker lists the column itself.
    return select
      ? select
          .split(',')
          .map((s) => s.trim().replace(/^_(.+)_value$/, '$1'))
          .filter(Boolean)
      : [];
  };

  // ── $batch ──
  if (/^\$batch$/i.test(path)) {
    if (method !== 'POST') return fail('A $batch request is a POST.');
    if (!cmd.body) return fail('The $batch request has no body.');
    let parts;
    try {
      parts = parseBatchRequest(cmd.body, contentType ?? '');
    } catch (e) {
      return fail(`The $batch body can't be read: ${e instanceof Error ? e.message : String(e)}`);
    }
    if (pref.unknown.length) warnings.push(`Left out Prefer: ${pref.unknown.join(', ')}.`);
    if (
      pref.spec.returnRepresentation ||
      pref.spec.maxpagesize != null ||
      prefer?.includes('annotations')
    )
      warnings.push(
        'Left out the batch-level Prefer preferences other than odata.continue-on-error: put them on the requests that need them.',
      );
    // DRS writes the batch's own OData / Accept / Content-Type headers.
    const outer = cmd.headers
      .filter((h) => !isCredentialHeader(h.name))
      .filter(
        (h) =>
          !/^(content-type|accept|odata-version|odata-maxversion|prefer)$/i.test(h.name) &&
          !BROWSER_HEADERS.has(h.name.toLowerCase()) &&
          !h.name.toLowerCase().startsWith('sec-') &&
          !(h.name.toLowerCase() === 'if-none-match' && h.value.trim() === 'null'),
      );
    if (outer.length)
      warnings.push(
        `Left out ${outer.map((h) => h.name).join(', ')} on the batch itself: DRS sends headers on each request in it.`,
      );
    const groups: BatchGroup[] = [];
    let n = 0;
    let renumbered = false;
    const dropped = new Set<string>();
    for (const g of parts) {
      const items = g.requests.map((r) => {
        n += 1;
        if (r.contentId && r.contentId !== String(n)) renumbered = true;
        const partHost = /^https?:\/\/([^/?#]+)/i.exec(r.url)?.[1]?.toLowerCase();
        if (partHost && partHost !== (host || ctx.host.toLowerCase()))
          warnings.push(
            `Request ${n} is for ${partHost}; it goes to the environment you're connected to.`,
          );
        const headers = Object.entries(r.headers)
          .filter(([k]) => {
            if (isCredentialHeader(k)) dropped.add(k);
            return !isCredentialHeader(k) && k !== 'content-id';
          })
          .map(([name, value]) => ({ name, value, enabled: true }));
        const caller = headers.find((h) => impersonationVar(h.name));
        if (caller) warnings.push(`Request ${n} runs as another user (${caller.name}).`);
        return newItem('custom', {
          method: r.method as 'GET' | 'POST' | 'PATCH' | 'PUT' | 'DELETE',
          url: r.url,
          body: r.body,
          headers,
        });
      });
      // Independent requests in a row share one group, as DRS shows them.
      const last = groups[groups.length - 1];
      if (!g.atomic && last && !last.atomic) last.items.push(...items);
      else groups.push(newGroup(g.atomic, items));
    }
    if (n === 0) return fail('The $batch body has no requests.');
    if (dropped.size) notes.push(`Left out ${[...dropped].join(', ')} from the batch's requests.`);
    if (renumbered)
      warnings.push(
        'DRS numbers Content-IDs 1, 2, 3 … in batch order; check any $n references against the new numbers.',
      );
    return done({ modeId: 'batch', groups, continueOnError: pref.continueOnError });
  }

  // ── Path: <name>[(<key>)][/<tail>] ──
  const seg = /^([A-Za-z_][A-Za-z0-9_.]*)(?:\(([^)]*)\))?(?:\/(.*))?$/.exec(safeDecode(path));
  if (!seg) return fail(`"${path}" isn't an entity set, row or operation.`);
  const [, name, key, tail = ''] = seg;
  const logical = name.includes('.') ? undefined : ctx.resolveEntitySet(name);

  if (!logical) {
    // An unbound operation: POST <Action>, GET <Function>(…).
    const opName = name.startsWith(NS) ? name.slice(NS.length) : name;
    if (tail || (key !== undefined && GUID_RE.test(key.trim())))
      return fail(`No table with the entity set name "${name}" in this environment.`);
    if (method === 'GET') {
      const fn = await ctx.findAction?.(opName).catch(() => undefined);
      if (fn?.kind === 'Function')
        return unsupported('functions', `Open Execute Function and pick ${opName}.`);
      return fail(`No table or function named "${name}" in this environment.`);
    }
    if (method !== 'POST')
      return fail(`No table with the entity set name "${name}" in this environment.`);
    if (key !== undefined) return fail(`An action is called without parentheses: POST /${opName}.`);
    return action(opName, null, null);
  }

  // A table: <set>, <set>(<id>), <set>(<id>)/<tail>.
  let recordId: string | null = null;
  if (key !== undefined) {
    if (GUID_RE.test(key.trim())) recordId = cleanGuid(key.trim());
    else if (key.includes('='))
      return unsupported(
        'rows addressed by an alternate key',
        'Open Update or Upsert and enter the key there.',
      );
    else return fail(`"${key}" isn't a row id (a GUID).`);
  }

  if (tail.startsWith(NS)) {
    const opName = tail.slice(NS.length);
    if (!IDENT_RE.test(opName)) return fail(`"${tail}" isn't an operation name.`);
    if (method === 'GET')
      return unsupported('bound functions', `Open Execute Function and pick ${opName}.`);
    if (method !== 'POST') return fail(`Call ${opName} with POST.`);
    return action(opName, logical, recordId);
  }

  if (tail && !/^\$count$/i.test(tail)) {
    if (/\/\$ref$/i.test(tail) || /^\$ref$/i.test(tail))
      return unsupported('associations', 'Use Associate or Disassociate.');
    if (!IDENT_RE.test(tail)) return fail(`"${tail}" isn't a column or navigation property.`);
    if (method === 'DELETE' && recordId) {
      const tbl = await ctx.loadTable?.(logical);
      if (tbl && !tbl.columns.some((c) => c.logicalName === tail))
        return fail(`${logical} has no column named "${tail}".`);
      noBody();
      return deleteRow({ kind: 'single-property', column: tail });
    }
    if (method === 'PUT')
      return unsupported('single-column updates (PUT)', 'Open Update and pick the column.');
    return unsupported(
      `${method} /${name}(…)/${tail}`,
      'Import opens table rows and queries, actions and $batch.',
    );
  }

  if (method === 'GET') {
    noBody();
    if (!recordId) {
      prefNotes('read');
      const r = await parseODataUrl(
        `${name}${tail ? `/${tail}` : ''}${query ? `?${query}` : ''}`,
        ctx.resolveEntitySet,
        ctx.loadTable,
      );
      warnings.push(...r.warnings.map((w) => w.message));
      if (!r.ok || !r.parsed || r.errors.length) {
        errors.push(...r.errors.map((e) => e.message));
        if (!errors.length) errors.push("The query couldn't be read.");
        return { ...summary, notes, warnings, errors };
      }
      return done({
        modeId: 'retrieve-multiple',
        parsed: r.parsed,
        prefer: { ...pref.spec, returnRepresentation: false },
        headers: readRows(),
      });
    }
    if (tail) return fail("$count counts a table's rows, not one row.");
    prefNotes('read');
    const r = await parseODataUrl(
      `${name}${query ? `?${query}` : ''}`,
      ctx.resolveEntitySet,
      ctx.loadTable,
    );
    warnings.push(...r.warnings.map((w) => w.message));
    if (!r.ok || !r.parsed || r.errors.length) {
      errors.push(...r.errors.map((e) => e.message));
      if (!errors.length) errors.push("The query couldn't be read.");
      return { ...summary, notes, warnings, errors };
    }
    noPageSize();
    const left = [...queryParams(query).keys()].filter((k) => k !== '$select' && k !== '$expand');
    if (left.length)
      warnings.push(`Left out ${left.join(', ')}: Retrieve Single sends only $select and $expand.`);
    return done({
      modeId: 'retrieve-single',
      table: logical,
      recordId,
      select: r.parsed.select,
      expand: r.parsed.expand,
      prefer: { ...pref.spec, returnRepresentation: false, maxpagesize: null },
      headers: readRows(),
    });
  }

  if (tail) return fail(`$count is only read with GET.`);

  if (method === 'POST') {
    if (recordId) return fail(`POST creates a row in /${name}, not in /${name}(${recordId}).`);
    prefNotes('write');
    const rawBody = jsonBody(false);
    if (rawBody === null) return { ...summary, notes, warnings, errors };
    if (cmd.body === undefined) warnings.push('The command has no body; Create opens with {}.');
    noPreconditions("Create doesn't send one");
    noPageSize();
    bypassNote();
    return done({
      modeId: 'create',
      table: logical,
      rawBody,
      prefer: { ...pref.spec, maxpagesize: null },
      returnSelect: writeQuery(),
      bypass,
      headers: rest,
    });
  }

  if (!recordId) return fail(`${method} needs a row: /${name}(<id>).`);

  if (method === 'PATCH') {
    prefNotes('write');
    const rawBody = jsonBody(true);
    if (rawBody === null) return { ...summary, notes, warnings, errors };
    const concurrency = concurrencyOf(ifMatch, ifNoneMatch);
    noPageSize();
    bypassNote();
    // If-Match makes it an update; without it (or with If-None-Match) PATCH
    // creates the row when it's missing — an upsert.
    const modeId =
      concurrency.kind === 'update-only' || concurrency.kind === 'etag' ? 'update' : 'upsert';
    return done({
      modeId,
      table: logical,
      recordId,
      rawBody,
      prefer: { ...pref.spec, maxpagesize: null },
      returnSelect: writeQuery(),
      concurrency,
      bypass,
      headers: rest,
    });
  }

  if (method === 'DELETE') {
    noBody();
    return deleteRow({ kind: 'whole-row' });
  }

  return unsupported('whole-row PUT', 'Use PATCH to update a row.');

  function deleteRow(scope: DeleteScope): CurlImportResult {
    prefNotes('none');
    if (ifNoneMatch) warnings.push("Left out If-None-Match: Delete doesn't use it.");
    if (query) warnings.push("Left out the query string: Delete doesn't send one.");
    bypassNote();
    // Delete offers None or an ETag: a DELETE of a missing row fails anyway,
    // so `If-Match: *` adds nothing there.
    const etag = ifMatch && ifMatch.trim() !== '*' ? ifMatch.trim() : null;
    if (ifMatch && !etag)
      notes.push('Left out If-Match: * — a DELETE of a missing row fails without it.');
    return done({
      modeId: 'delete',
      table: logical!,
      recordId: recordId!,
      scope,
      concurrency: etag ? { kind: 'etag', etag } : { kind: 'none' },
      bypass,
      headers: rest,
    });
  }

  async function action(
    opName: string,
    boundTable: string | null,
    boundId: string | null,
  ): Promise<CurlImportResult> {
    if (!IDENT_RE.test(opName)) return fail(`"${opName}" isn't an action name.`);
    const found = await ctx.findAction?.(opName).catch(() => undefined);
    if (!found || found.kind !== 'Action')
      return fail(`No action named ${opName} in this environment.`);
    const b = found.binding;
    if (boundTable) {
      if (b.kind === 'unbound' || b.entityType !== boundTable)
        return fail(`${opName} isn't bound to ${boundTable}.`);
      if (b.kind === 'entity' && !boundId)
        return fail(`${opName} is bound to one row: /${name}(<id>)/${NS}${opName}.`);
      if (b.kind === 'collection' && boundId)
        return fail(`${opName} is bound to the table, not a row.`);
    } else if (b.kind !== 'unbound') {
      return fail(
        `${opName} is bound to ${b.entityType}; call it on /<${b.entityType} set>…/${NS}${opName}.`,
      );
    }
    prefNotes('none');
    noQuery();
    noPreconditions("actions don't use one");
    const rawBody = jsonBody(false);
    if (rawBody === null) return { ...summary, notes, warnings, errors };
    // The form keeps parameters in its own shapes (a row reference as
    // { id, entityType }, …); read the wire JSON back into them so Execute
    // sends what the command sent.
    for (const t of new Set(found.parameters.map((p) => p.entityType).filter(Boolean)))
      await ctx.loadTable?.(t!).catch(() => undefined);
    const paramValues: Record<string, unknown> = {};
    for (const [k, v] of Object.entries(JSON.parse(rawBody) as Record<string, unknown>)) {
      const param = found.parameters.find((p) => p.name === k);
      if (!param) {
        errors.push(`${found.name} has no parameter named "${k}".`);
        continue;
      }
      const r = decodeParamFromBody(param, v, ctx.resolveEntitySet);
      if ('error' in r) errors.push(r.error);
      else paramValues[k] = r.value;
    }
    if (errors.length) return { ...summary, notes, warnings, errors };
    const modeId =
      found.source === 'custom-api'
        ? 'exec-customapi'
        : found.source === 'custom-action'
          ? 'exec-customaction'
          : 'exec-action';
    return done({
      modeId,
      actionName: found.name,
      boundRecordId: boundId,
      paramValues,
      headers: kept,
    });
  }

  function noQuery() {
    if (query) warnings.push("Left out the query string: actions don't use one.");
  }
}
