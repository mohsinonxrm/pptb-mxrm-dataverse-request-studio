// OData → RetrieveMultipleState parser.
//
// Takes a Dataverse Web API URL (relative or absolute) and produces the
// filter / select / expand / orderby / top / count / apply state shape DRS
// uses. Lets users paste a known-good query (from MS Learn docs, DRB,
// F12-console testing, saved snippets) and get a builder-driven view of
// it for editing.
//
// Scope:
//   ✅ entity set name → table logical name (via entity list lookup)
//   ✅ $select — splits on commas, decodes `_<x>_value` back to bare lookup
//   ✅ $top / $skip / $count
//   ✅ $orderby — multi-column, asc/desc
//   ✅ $filter — full recursive-descent parser supporting:
//       • comparison: eq/ne/gt/ge/lt/le, literals + same-table col-vs-col
//       • boolean: and, or, not (root + grouped subexpressions)
//       • OData string fns: contains/startswith/endswith, with optional `not`
//       • Dataverse query fns: Microsoft.Dynamics.CRM.* with all arities
//       • lambdas: <nav>/any(<alias>: <inner>) and /all(...)
//       • nested lambdas with outer-alias prefix on inner path
//   ✅ $expand — recursive, with nested $select / $filter / $orderby / $top / $expand
//   ✅ $apply — filter() / groupby((cols), aggregate(<aggs>)) / aggregate(<aggs>)
//
// Out of scope (warns instead):
//   ⚠️ `not (<comparison>)` — DRS's filter model has no per-rule NOT for
//      comparison rules. We drop the `not` and surface a warning so the user
//      knows the parse isn't lossless. (Encoder still produces the right
//      result for the non-negated form.)
//   ⚠️ `$apply` with `bottomcount`/`topcount` or compute() — uncommon, not
//      currently modeled.
//   ⚠️ `$search`, `$format`, `$skiptoken` — Dataverse doesn't support these
//      and DRS doesn't model them. Skipped silently.

import type { TableMeta, NavProperty } from '../mock/metadata';
import { findTable } from '../mock/metadata';
import type {
  FilterGroup, FilterNode, FilterRule, FilterFunctionNode, FilterLambdaNode,
} from '../editors/filter/filterTree';
import { newId } from '../editors/filter/filterTree';
import { findOperator, OPERATORS } from '../editors/filter/operators';
import type { ExpandSpec } from '../editors/ExpandEditor';
import { makeExpandSpec } from '../editors/ExpandEditor';
import type { OrderbySpec } from '../editors/OrderbyEditor';
import type { ApplySpec, ApplyAgg, AggFn } from '../editors/ApplyEditor';

// ── Public surface ──────────────────────────────────────────────────────

export interface ParsedRequest {
  /** Logical name (e.g. 'account'), resolved from entitySetName. Empty when the entity is unknown. */
  table: string;
  /** Entity-set name as it appeared in the URL (e.g. 'accounts'). */
  entitySet: string;
  select: string[];
  filter: FilterGroup;
  orderby: OrderbySpec[];
  top: number | null;
  countOn: boolean;
  expand: ExpandSpec[];
  apply?: ApplySpec;
}

/**
 * A single parse-time message. `message` is short (one sentence) and
 * actionable. `learnMoreUrl` optionally points at the relevant MS Learn
 * page so the user can verify what DRS supports without us having to
 * enumerate every value in the message itself.
 */
export interface ParseIssue {
  message: string;
  /** Optional MS Learn deep link for the user to verify our claim. */
  learnMoreUrl?: string;
}

export interface ParseResult {
  ok: boolean;
  parsed?: ParsedRequest;
  /** Lossy-but-legal constructs DRS can still represent. Don't block apply. */
  warnings: ParseIssue[];
  /** Things DRS can't faithfully represent. Block apply. */
  errors: ParseIssue[];
}

// MS Learn deep links for common error categories — used as `learnMoreUrl`.
const DOCS = {
  queryFunctions: 'https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/reference/queryfunctions',
  filterRows:     'https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/query/filter-rows',
  aggregateData:  'https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/query/aggregate-data',
  queryOptions:   'https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/query-data-web-api',
  joinTables:     'https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/query/join-tables',
  attrMetadata:   'https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/use-web-api-functions#address-lookup-and-customer-fields-in-functions',
};

/** Internal — throw from any sub-parser to carry a learn-more URL up through
 *  the top-level catch. */
class ParseFailure extends Error {
  constructor(message: string, public learnMoreUrl?: string) {
    super(message);
  }
}

/** Short constructor for ParseIssue. */
const iss = (message: string, learnMoreUrl?: string): ParseIssue =>
  ({ message, learnMoreUrl });

/** Prefix the message of an existing issue, preserving its learnMoreUrl. */
const prefixIss = (prefix: string, i: ParseIssue): ParseIssue =>
  ({ message: `${prefix}: ${i.message}`, learnMoreUrl: i.learnMoreUrl });

/**
 * Resolve an entity-set name (e.g. 'accounts') to a logical name (e.g. 'account').
 * Looks through the live registry of TableMetas; returns the first match.
 * Callers pass a getter to keep this module pure.
 */
export type EntitySetResolver = (entitySetName: string) => string | undefined;

/**
 * Async loader for full TableMeta — pulls columns + relationships from
 * Dataverse if not already cached. Callers wire this to
 * `metadataProvider.getTable(logical)`.
 */
export type LoadTable = (logical: string) => Promise<TableMeta | undefined>;

// ── Static allow-lists for validation ─────────────────────────────────

/** OData / Dataverse query options DRS understands. Anything else is rejected. */
const KNOWN_QUERY_OPTIONS = new Set([
  '$select', '$filter', '$orderby', '$top', '$skip', '$count', '$expand', '$apply',
]);

/** Every Dataverse query function defined in the operator registry. */
const KNOWN_DV_FUNCTIONS = new Set(
  OPERATORS.filter(o => o.kind.startsWith('dv-fn-')).map(o => o.id),
);

/** Aggregate functions DRS's ApplyEditor can model. `countdistinct` is
 *  valid OData but not modeled by DRS — pasted requests using it must be
 *  rewritten manually. */
const KNOWN_AGG_FNS = new Set(['sum', 'average', 'min', 'max', '$count']);

// ── Top-level entry point ───────────────────────────────────────────────

/**
 * Parse → validate → return. The pipeline is:
 *
 *   1. Lex + parse  — syntactic shape only. Throws on syntax errors.
 *   2. Static check — entity set known? query options known? DV functions
 *                     in registry? aggregate functions modelled?
 *   3. Semantic check (async) — columns in $select / $orderby /
 *                     $expand($select) / groupby / aggregate exist on the
 *                     right entity? nav properties in $expand exist?
 *                     Requires metadata loading for every involved entity.
 *
 * Any errors from (2) or (3) block the Apply button. Warnings remain for
 * legitimately lossy but legal patterns (e.g. `not Microsoft.Dynamics.CRM.X`
 * → drop the `not`, function preserved — query still works).
 */
export async function parseODataUrl(
  rawUrl: string,
  resolveEntitySet: EntitySetResolver,
  loadTable?: LoadTable,
): Promise<ParseResult> {
  const warnings: ParseIssue[] = [];
  const errors: ParseIssue[] = [];
  try {
    const { entitySet, query } = splitUrl(rawUrl);
    if (!entitySet) {
      errors.push(iss('Could not extract an entity-set name from the URL.', DOCS.queryOptions));
      return { ok: false, warnings, errors };
    }

    // ── Static: entity set resolution ──
    const table = resolveEntitySet(entitySet) ?? '';
    if (!table) {
      errors.push(iss(
        `Entity set '${entitySet}' is not a known entity in this environment.`,
        DOCS.queryOptions,
      ));
      return { ok: false, warnings, errors };
    }
    const tbl = findTable(table);

    const opts = splitQueryOptions(query);

    // ── Static: unknown query options ──
    for (const k of Object.keys(opts)) {
      if (!KNOWN_QUERY_OPTIONS.has(k)) {
        errors.push(iss(
          `Unknown query option \`${k}\` — not part of DRS's supported set.`,
          DOCS.queryOptions,
        ));
      }
    }
    if (errors.length) return { ok: false, warnings, errors };

    // ── Scalars ──
    const top = opts.$top ? parseIntOrNull(opts.$top) : null;
    const countOn = opts.$count === 'true';

    // ── $select ──
    const select = opts.$select ? parseSelect(opts.$select, tbl) : [];

    // ── $orderby ──
    const orderby = opts.$orderby ? parseOrderby(opts.$orderby, tbl) : [];

    // ── $filter ── (collects its own errors for unknown DV functions)
    let filter: FilterGroup = emptyFilterGroup();
    if (opts.$filter) {
      const r = parseFilter(opts.$filter, tbl);
      filter = r.group;
      warnings.push(...r.warnings);
      errors.push(...r.errors);
    }

    // ── $expand ──
    let expand: ExpandSpec[] = [];
    if (opts.$expand) {
      const r = parseExpand(opts.$expand, tbl);
      expand = r.items;
      warnings.push(...r.warnings);
      errors.push(...r.errors);
    }

    // ── $apply ──
    let apply: ApplySpec | undefined;
    if (opts.$apply) {
      const r = parseApply(opts.$apply, tbl);
      apply = r.spec;
      warnings.push(...r.warnings);
      errors.push(...r.errors);
    }

    // If any sub-parser flagged a hard error, bail before semantic validation.
    if (errors.length) return { ok: false, warnings, errors };

    const parsed: ParsedRequest = { table, entitySet, select, filter, orderby, top, countOn, expand, apply };

    // ── Semantic validation (async) ──
    // Loads metadata for every involved entity and checks column / nav
    // references exist. If loadTable wasn't supplied (e.g. unit tests),
    // we skip this phase entirely — orphans surface in the editors later.
    if (loadTable) {
      await validateAgainstMetadata(parsed, loadTable, errors, warnings);
    }

    return {
      ok: errors.length === 0,
      parsed: errors.length === 0 ? parsed : undefined,
      warnings,
      errors,
    };
  } catch (e) {
    if (e instanceof ParseFailure) {
      errors.push(iss(e.message, e.learnMoreUrl));
    } else {
      errors.push(iss(e instanceof Error ? e.message : String(e), DOCS.filterRows));
    }
    return { ok: false, warnings, errors };
  }
}

// ── Semantic validator (uses async-loaded metadata) ─────────────────────

/**
 * Wrap loadTable with a timeout safety net so a hung metadata fetch can't
 * leave the dialog spinner stuck forever. Returns undefined on timeout;
 * caller emits a warning and skips that entity's validation.
 */
async function loadTableWithTimeout(
  loadTable: LoadTable,
  logical: string,
  timeoutMs = 15000,
): Promise<TableMeta | undefined> {
  return Promise.race([
    loadTable(logical),
    new Promise<undefined>(resolve => setTimeout(() => resolve(undefined), timeoutMs)),
  ]);
}

async function validateAgainstMetadata(
  parsed: ParsedRequest,
  loadTable: LoadTable,
  errors: ParseIssue[],
  warnings: ParseIssue[],
): Promise<void> {
  const rootTbl = await loadTableWithTimeout(loadTable, parsed.table);
  if (!rootTbl) {
    warnings.push(iss(
      `Metadata for \`${parsed.table}\` could not be loaded (or timed out) — column validation skipped.`,
    ));
    return;
  }

  // ── $select on root ──
  for (const c of parsed.select) {
    if (!columnExistsOnTable(rootTbl, c)) {
      errors.push(iss(
        `$select: column \`${c}\` does not exist on \`${parsed.table}\`.`,
        DOCS.attrMetadata,
      ));
    }
  }

  // ── $orderby on root ──
  for (const o of parsed.orderby) {
    if (!columnExistsOnTable(rootTbl, o.col)) {
      errors.push(iss(
        `$orderby: column \`${o.col}\` does not exist on \`${parsed.table}\`.`,
        DOCS.attrMetadata,
      ));
    }
  }

  // ── $filter rule columns (recursive — walks lambdas + nested groups) ──
  await validateFilterColumns(parsed.filter, rootTbl, loadTable, errors, warnings, undefined);

  // ── $expand tree (recursive) ──
  for (const e of parsed.expand) {
    await validateExpandTree(e, rootTbl, loadTable, errors, warnings);
  }

  // ── $apply ──
  if (parsed.apply?.enabled) {
    for (const g of parsed.apply.groupby) {
      const segs = g.split('/');
      if (segs.length === 1) {
        if (!columnExistsOnTable(rootTbl, segs[0])) {
          errors.push(iss(
            `$apply/groupby: column \`${segs[0]}\` does not exist on \`${parsed.table}\`.`,
            DOCS.aggregateData,
          ));
        }
      } else {
        const firstHop = segs[0];
        const nav = rootTbl.navigationProperties.find(n => n.name === firstHop);
        if (!nav && !columnExistsOnTable(rootTbl, firstHop)) {
          errors.push(iss(
            `$apply/groupby: \`${firstHop}\` is not a column or navigation on \`${parsed.table}\`.`,
            DOCS.aggregateData,
          ));
        }
      }
    }
    for (const a of parsed.apply.aggregates) {
      if (a.fn !== '$count' && !columnExistsOnTable(rootTbl, a.col)) {
        errors.push(iss(
          `$apply/aggregate: column \`${a.col}\` does not exist on \`${parsed.table}\`.`,
          DOCS.aggregateData,
        ));
      }
    }
    // Also validate the prefilter (uses same root table).
    await validateFilterColumns(parsed.apply.prefilter, rootTbl, loadTable, errors, warnings, undefined);
  }
}

/**
 * Walk the parsed filter tree, validating each rule's column reference
 * against the table it logically belongs to. Lambda predicates recurse
 * into the lambda's target table.
 *
 * Skipped (accepted as-is):
 *   • Multi-segment nav-paths (`primarycontactid/fullname`) — would
 *     require lazy-loading every intermediate target table.
 *   • Lambda nav references when the target table can't be loaded.
 */
async function validateFilterColumns(
  group: FilterGroup,
  ownerTable: TableMeta,
  loadTable: LoadTable,
  errors: ParseIssue[],
  warnings: ParseIssue[],
  lambdaAlias: string | undefined,
): Promise<void> {
  for (const node of group.rules) {
    if (node.type === 'group') {
      await validateFilterColumns(node, ownerTable, loadTable, errors, warnings, lambdaAlias);
    } else if (node.type === 'rule' || node.type === 'function') {
      let col = node.col;
      if (!col) continue;
      // Strip lambda alias prefix
      if (lambdaAlias && col.startsWith(lambdaAlias + '/')) {
        col = col.slice(lambdaAlias.length + 1);
      }
      // Multi-segment nav-path (e.g. `primarycontactid/abc_salesstage`):
      // walk the N:1 hops, loading each related entity — which both validates
      // the leaf AND warms the related metadata so the $filter encoder can
      // resolve the leaf's type (and avoid string-quoting a numeric/boolean
      // OptionSet — issue #33). Reported as a warning, not an error, because
      // a hop may legitimately fail to load within the timeout.
      if (col.includes('/')) {
        const segs = col.split('/');
        let cursor: TableMeta | undefined = ownerTable;
        for (let i = 0; i < segs.length - 1; i++) {
          if (!cursor) break;
          const nav: NavProperty | undefined = cursor.navigationProperties.find(
            (n: NavProperty) => n.name === segs[i] && n.cardinality === 'ManyToOne',
          );
          if (!nav) { cursor = undefined; break; }
          cursor = await loadTableWithTimeout(loadTable, nav.targetEntity);
        }
        if (cursor) {
          const leaf = segs[segs.length - 1];
          if (!columnExistsOnTable(cursor, leaf)) {
            warnings.push(iss(
              `$filter: column \`${leaf}\` not found on \`${cursor.logicalName}\` (via \`${col}\`).`,
              DOCS.filterRows,
            ));
          }
        }
        continue;
      }
      if (!columnExistsOnTable(ownerTable, col)) {
        errors.push(iss(
          `$filter: column \`${col}\` does not exist on \`${ownerTable.logicalName}\`.`,
          DOCS.filterRows,
        ));
      }
    } else if (node.type === 'lambda') {
      const nav = ownerTable.navigationProperties.find(n => n.name === node.nav);
      if (!nav) {
        errors.push(iss(
          `$filter lambda: nav \`${node.nav}\` does not exist on \`${ownerTable.logicalName}\`.`,
          DOCS.filterRows,
        ));
        continue;
      }
      const targetTbl = await loadTableWithTimeout(loadTable, nav.targetEntity);
      if (!targetTbl) {
        warnings.push(iss(
          `Metadata for \`${nav.targetEntity}\` (lambda target via \`${node.nav}\`) could not be loaded — inner filter not column-validated.`,
        ));
        continue;
      }
      await validateFilterColumns(node.inner, targetTbl, loadTable, errors, warnings, node.alias);
    }
  }
}

async function validateExpandTree(
  spec: ExpandSpec,
  parentTbl: TableMeta,
  loadTable: LoadTable,
  errors: ParseIssue[],
  warnings: ParseIssue[],
): Promise<void> {
  const nav = parentTbl.navigationProperties.find(n => n.name === spec.nav);
  if (!nav) {
    errors.push(iss(
      `$expand: navigation property \`${spec.nav}\` does not exist on \`${parentTbl.logicalName}\`.`,
      DOCS.joinTables,
    ));
    return;
  }
  const targetTbl = await loadTableWithTimeout(loadTable, nav.targetEntity);
  if (!targetTbl) {
    warnings.push(iss(
      `Metadata for \`${nav.targetEntity}\` (target of \`${spec.nav}\`) could not be loaded — inner $select not validated.`,
    ));
    return;
  }
  for (const c of spec.select) {
    if (!columnExistsOnTable(targetTbl, c)) {
      errors.push(iss(
        `$expand=${spec.nav}: column \`${c}\` does not exist on \`${nav.targetEntity}\`.`,
        DOCS.attrMetadata,
      ));
    }
  }
  for (const o of spec.orderby) {
    if (!columnExistsOnTable(targetTbl, o.col)) {
      errors.push(iss(
        `$expand=${spec.nav}/$orderby: column \`${o.col}\` does not exist on \`${nav.targetEntity}\`.`,
        DOCS.attrMetadata,
      ));
    }
  }
  if (spec.filter) {
    await validateFilterColumns(spec.filter, targetTbl, loadTable, errors, warnings, undefined);
  }
  if (spec.nestedExpand) {
    for (const ne of spec.nestedExpand) {
      await validateExpandTree(ne, targetTbl, loadTable, errors, warnings);
    }
  }
}

/**
 * True if `name` matches a real column on the table — either by logical
 * name or by the pre-computed `oDataName` (which covers the `_<x>_value`
 * wire form for lookups).
 */
function columnExistsOnTable(table: TableMeta, name: string): boolean {
  return table.columns.some(c => c.logicalName === name || c.oDataName === name);
}

// ── URL splitting helpers ───────────────────────────────────────────────

function splitUrl(rawUrl: string): { entitySet: string; query: string } {
  let s = rawUrl.trim();
  // Strip absolute prefix (https://<host>/api/data/vX.Y/) or any leading slash.
  s = s.replace(/^https?:\/\/[^/]+/, '');
  s = s.replace(/^\/?api\/data\/v\d+(\.\d+)?\//, '');
  s = s.replace(/^\//, '');
  // Decode percent-encoding — pasted URLs are often url-encoded.
  // Decode once cautiously; some test runners log already-decoded URLs.
  if (/%[0-9A-Fa-f]{2}/.test(s)) {
    try { s = decodeURIComponent(s); } catch { /* keep original */ }
  }
  const qIdx = s.indexOf('?');
  if (qIdx < 0) return { entitySet: stripKeyPart(s), query: '' };
  const entitySet = stripKeyPart(s.slice(0, qIdx));
  const query = s.slice(qIdx + 1);
  return { entitySet, query };
}

/** `accounts(<guid>)` → `accounts` — keys aren't relevant for retrieve-multiple. */
function stripKeyPart(p: string): string {
  const i = p.indexOf('(');
  return i >= 0 ? p.slice(0, i) : p;
}

/**
 * Split `$opt=...&$opt2=...` into a record, respecting `&` outside parens.
 * (Inside `$expand=primarycontactid($select=fullname;$filter=name eq 'x')`
 * the inner `&` would have been encoded — but DRS-generated URLs use `;`
 * internally. We still scan for balanced parens to be safe.)
 */
function splitQueryOptions(query: string): Record<string, string> {
  if (!query) return {};
  const out: Record<string, string> = {};
  const parts = topLevelSplit(query, '&');
  for (const part of parts) {
    const eq = part.indexOf('=');
    if (eq < 0) continue;
    const k = part.slice(0, eq);
    const v = part.slice(eq + 1);
    out[k] = v;
  }
  return out;
}

/**
 * Split a string on the given delimiter, but only at the OUTERMOST nesting
 * level. Respects parens and single-quoted strings.
 *
 * Used everywhere we need to chop a comma-separated / semicolon-separated /
 * slash-separated list whose elements themselves contain parens/quotes:
 *   • query options ('&' top-level)
 *   • $expand sibling items (',' between expands)
 *   • inner expand options (';' inside expand parens)
 *   • $apply pipeline stages ('/' between filter/groupby/aggregate)
 */
export function topLevelSplit(s: string, delim: string): string[] {
  const out: string[] = [];
  let depth = 0;
  let bracketDepth = 0;
  let inString = false;
  let buf = '';
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (inString) {
      buf += c;
      if (c === "'") {
        if (s[i + 1] === "'") { buf += "'"; i++; continue; }
        inString = false;
      }
      continue;
    }
    if (c === "'") { inString = true; buf += c; continue; }
    if (c === '(') { depth++; buf += c; continue; }
    if (c === ')') { depth--; buf += c; continue; }
    if (c === '[') { bracketDepth++; buf += c; continue; }
    if (c === ']') { bracketDepth--; buf += c; continue; }
    if (c === delim && depth === 0 && bracketDepth === 0) {
      out.push(buf);
      buf = '';
      continue;
    }
    buf += c;
  }
  if (buf.length) out.push(buf);
  return out;
}

// ── $select parser ──────────────────────────────────────────────────────

function parseSelect(s: string, table?: TableMeta): string[] {
  return s.split(',').map(c => c.trim()).filter(Boolean)
    .map(col => decodeLookupValueForm(col, table));
}

/**
 * Decode the `_<x>_value` lookup-column wire form back into the bare logical
 * name DRS stores — IF the target table actually has `<x>` as a Lookup /
 * Customer / Owner column. Otherwise preserve the wire form unchanged so
 * the encoder doesn't blindly translate a nonexistent attribute (which
 * would turn Dataverse's lenient silent-skip into a hard 400).
 *
 *   _primarycontactid_value on `account`  → 'primarycontactid' (col exists, is Lookup)
 *   _accountid_value        on `contact`  → '_accountid_value' (col doesn't exist — preserved)
 *   accountid               on `account`  → 'accountid'        (already bare)
 *
 * When `table` is omitted (metadata not yet loaded, common for inner
 * `$expand` paths whose target entity hasn't been lazy-fetched yet), we
 * CONSERVATIVELY PRESERVE the wire form. Rationale:
 *   • If `_X_value` is a real lookup, the encoder will re-emit it as
 *     `_X_value` regardless (attrRefByName returns the original logicalName
 *     when col isn't found). Faithful.
 *   • If `_X_value` is an unknown column, preserving avoids decoding to a
 *     bare `X` that the encoder also can't re-wrap → which would produce
 *     a different (DRS-side) 400 from the (legitimate, server-side) 400
 *     the user is actually testing.
 *   • The cost of preserving is purely cosmetic: the SelectEditor will
 *     render the value as an "(unknown)" orphan until metadata loads.
 *     The orphan banner explains what happened.
 */
function decodeLookupValueForm(col: string, table?: TableMeta): string {
  if (!col.startsWith('_') || !col.endsWith('_value') || col.length <= '_value'.length + 1) {
    return col;
  }
  const bare = col.slice(1, -'_value'.length);
  if (!table) return col; // No metadata available — preserve wire form (see comment above).
  const meta = table.columns.find(c => c.logicalName === bare);
  if (meta && (meta.attributeType === 'Lookup' || meta.attributeType === 'Customer' || meta.attributeType === 'Owner')) {
    return bare;
  }
  // Column doesn't exist or isn't a lookup → preserve `_value` form so the
  // round-trip is faithful. Encoder's `attrRefByName` falls through to the
  // logicalName as-is when no matching column is found.
  return col;
}

// ── $orderby parser ─────────────────────────────────────────────────────

function parseOrderby(s: string, table?: TableMeta): OrderbySpec[] {
  return s.split(',').map(c => c.trim()).filter(Boolean).map(seg => {
    // `<col> asc|desc?`
    const m = seg.match(/^(\S+)(?:\s+(asc|desc))?$/i);
    const id = newId('ord');
    if (!m) return { id, col: seg, dir: 'asc' as const };
    return {
      id,
      col: decodeLookupValueForm(m[1], table),
      dir: (m[2]?.toLowerCase() === 'desc' ? 'desc' : 'asc') as 'asc' | 'desc',
    };
  });
}

// ── $filter parser ──────────────────────────────────────────────────────
//
// Tokenize → recursive descent → FilterGroup tree.

interface Tok {
  kind: 'ident' | 'string' | 'number' | 'lparen' | 'rparen' | 'comma' | 'colon' | 'bracket';
  value: string;
}

function tokenize(input: string): Tok[] {
  const out: Tok[] = [];
  let i = 0;
  while (i < input.length) {
    const c = input[i];
    if (c === ' ' || c === '\t' || c === '\n' || c === '\r') { i++; continue; }
    if (c === '(') { out.push({ kind: 'lparen', value: '(' }); i++; continue; }
    if (c === ')') { out.push({ kind: 'rparen', value: ')' }); i++; continue; }
    if (c === ',') { out.push({ kind: 'comma', value: ',' }); i++; continue; }
    if (c === ':') {
      // A standalone colon means either:
      //   • lambda alias separator: `any(c: <expr>)` — `:` here is NOT
      //     followed by a digit, so emit the colon token.
      //   • inside an ISO datetime literal: `2024-01-01T00:00:00Z` — the
      //     `:` is between digits. The previous ident accumulator already
      //     handles this case (we never enter the dispatch loop mid-ident),
      //     but if a value starts with `:` followed by a digit, fall through
      //     to treat the whole run as one ident token. Belt-and-suspenders.
      const next = input[i + 1];
      if (next && /\d/.test(next)) {
        // Fall through to the ident accumulator.
      } else {
        out.push({ kind: 'colon', value: ':' }); i++; continue;
      }
    }
    if (c === "'") {
      // OData string literal — '' is an escaped single quote.
      let s = ''; i++;
      while (i < input.length) {
        if (input[i] === "'" && input[i + 1] === "'") { s += "'"; i += 2; continue; }
        if (input[i] === "'") { i++; break; }
        s += input[i++];
      }
      out.push({ kind: 'string', value: s });
      continue;
    }
    if (c === '[') {
      // PropertyValues=[...] — capture verbatim, JSON-decode at use site.
      let depth = 0; let s = '';
      while (i < input.length) {
        const ch = input[i];
        s += ch;
        if (ch === '[') depth++;
        else if (ch === ']') { depth--; if (depth === 0) { i++; break; } }
        i++;
      }
      out.push({ kind: 'bracket', value: s });
      continue;
    }
    // Identifier (possibly dotted/slashed/dashed/dot-namespaced):
    //   primarycontactid          col
    //   c/firstname               lambda-scoped col
    //   primarycontactid/fullname nav-path col
    //   Microsoft.Dynamics.CRM.X  DV fn ident
    //   true / false / null       literals (treated as idents)
    //   2024-01-01T00:00:00Z      datetime literal (treated as ident)
    //   2024-01-01T00:00:00+05:30 datetime with offset
    //   3.14, -42                 numbers (treated as idents until we type-check)
    //
    // The `:` character is special. It can mean:
    //   (a) lambda alias separator `c: <expr>` — `:` is followed by a
    //       letter or whitespace.
    //   (b) part of an ISO datetime — `:` is followed by a digit
    //       (`T00:00:00` is part of `2024-01-01T00:00:00Z`).
    // We include `:` in the ident ONLY when the next character is a digit.
    // Same rule for the offset-introducer `+` / `-` after a `:dd` pair —
    // but those are already non-break chars, so they handle themselves.
    let s = '';
    while (i < input.length) {
      const ch = input[i];
      if (' \t\n\r(),[\''.includes(ch)) break;
      if (ch === ':') {
        const next = input[i + 1];
        if (!next || !/\d/.test(next)) break;
      }
      s += ch;
      i++;
    }
    if (!s) { i++; continue; }
    // Classify number vs ident — pure number/decimal/scientific.
    if (/^-?\d[\d.eE+\-]*$/.test(s)) {
      out.push({ kind: 'number', value: s });
    } else {
      out.push({ kind: 'ident', value: s });
    }
  }
  return out;
}

interface ParseFilterCtx {
  toks: Tok[];
  pos: number;
  warnings: ParseIssue[];
  errors: ParseIssue[];
  /** Alias of the enclosing lambda, if we're recursing inside one. */
  outerAlias?: string;
}

const COMPARISON_OPS = new Set(['eq', 'ne', 'gt', 'ge', 'lt', 'le']);
const STRING_FN_NAMES = new Set(['contains', 'startswith', 'endswith']);
const BOOL_LITERALS = new Set(['true', 'false']);

function emptyFilterGroup(): FilterGroup {
  return { id: newId('root'), type: 'group', combinator: 'and', rules: [] };
}

function parseFilter(input: string, _table: TableMeta | undefined): { group: FilterGroup; warnings: ParseIssue[]; errors: ParseIssue[] } {
  const warnings: ParseIssue[] = [];
  const errors: ParseIssue[] = [];
  const ctx: ParseFilterCtx = { toks: tokenize(input), pos: 0, warnings, errors };
  if (ctx.toks.length === 0) return { group: emptyFilterGroup(), warnings, errors };
  try {
    const node = parseOrExpr(ctx);
    // Wrap a single-node result in a root group so the FilterEditor sees a
    // proper tree.
    const group: FilterGroup = node.type === 'group'
      ? node
      : { id: newId('root'), type: 'group', combinator: 'and', rules: [node] };
    return { group, warnings, errors };
  } catch (e) {
    // Convert thrown ParseFailure (or generic Error) into an entry in errors
    // rather than propagating up — gives the user every detected issue at
    // once instead of bailing at the first one.
    if (e instanceof ParseFailure) {
      errors.push(iss(e.message, e.learnMoreUrl));
    } else {
      errors.push(iss(e instanceof Error ? e.message : String(e), DOCS.filterRows));
    }
    return { group: emptyFilterGroup(), warnings, errors };
  }
}

/** orExpr := andExpr ('or' andExpr)* */
function parseOrExpr(ctx: ParseFilterCtx): FilterNode {
  const left = parseAndExpr(ctx);
  if (!isKeywordAhead(ctx, 'or')) return left;
  const children: FilterNode[] = [left];
  while (isKeywordAhead(ctx, 'or')) {
    consumeKeyword(ctx, 'or');
    children.push(parseAndExpr(ctx));
  }
  return {
    id: newId('g'),
    type: 'group',
    combinator: 'or',
    rules: children.flatMap(c =>
      (c.type === 'group' && c.combinator === 'or') ? c.rules : [c],
    ),
  };
}

/** andExpr := unaryExpr ('and' unaryExpr)* */
function parseAndExpr(ctx: ParseFilterCtx): FilterNode {
  const left = parseUnaryExpr(ctx);
  if (!isKeywordAhead(ctx, 'and')) return left;
  const children: FilterNode[] = [left];
  while (isKeywordAhead(ctx, 'and')) {
    consumeKeyword(ctx, 'and');
    children.push(parseUnaryExpr(ctx));
  }
  return {
    id: newId('g'),
    type: 'group',
    combinator: 'and',
    rules: children.flatMap(c =>
      (c.type === 'group' && c.combinator === 'and') ? c.rules : [c],
    ),
  };
}

/** unaryExpr := 'not' unaryExpr | primary */
function parseUnaryExpr(ctx: ParseFilterCtx): FilterNode {
  if (isKeywordAhead(ctx, 'not')) {
    consumeKeyword(ctx, 'not');
    const inner = parseUnaryExpr(ctx);
    return applyNot(inner, ctx);
  }
  return parsePrimary(ctx);
}

/**
 * Apply a `not` to the parsed inner node. The mapping mirrors what DRS's
 * filter model now supports (empirically locked):
 *
 *   • `not <contains/startswith/endswith>(…)` → `rule.negated = true`
 *     (test 3.4–3.6 → 200; legacy per-rule toggle).
 *   • `not <plain comparison>` → wrap the single rule in a `negated`
 *     group of one — same wire shape as `not (<comparison>)`, which is
 *     what Dataverse expects (test 3.3 / 3.7 → 200).
 *   • `not (<group>)` → set `group.negated = true` (tests G.1–G.7 → 200).
 *   • `not <lambda>` → set `lambda.negated = true` (test G.8 → 200).
 *   • `not <Microsoft.Dynamics.CRM.*>(…)` → DROP the `not`, warn loudly.
 *     Even `not (<dv-fn>)` is rejected with 405 (test G.9). Use the
 *     explicit `Not*` sibling fn instead.
 */
function applyNot(node: FilterNode, ctx: ParseFilterCtx): FilterNode {
  if (node.type === 'rule') {
    const op = findOperator(node.op);
    if (op && op.kind === 'odata-fn') {
      return { ...node, negated: true };
    }
    // Wrap the comparison in a `not (…)` group of one — encoder emits
    // exactly `not (<col> <op> <val>)` which Dataverse accepts.
    return {
      id: newId('g'),
      type: 'group',
      combinator: 'and',
      rules: [node],
      negated: true,
    };
  }
  if (node.type === 'group') {
    // Stacking `not` flips it back off rather than double-wrapping.
    return { ...node, negated: !node.negated };
  }
  if (node.type === 'lambda') {
    return { ...node, negated: !node.negated };
  }
  if (node.type === 'function') {
    ctx.warnings.push(iss(
      `\`not Microsoft.Dynamics.CRM.${node.op}(…)\` is rejected by Dataverse — dropped the \`not\`, function preserved. Use the explicit \`Not*\` sibling instead.`,
      DOCS.queryFunctions,
    ));
    return node;
  }
  return node;
}

/** primary := '(' expr ')' | dvFn | strFn | lambda | comparison */
function parsePrimary(ctx: ParseFilterCtx): FilterNode {
  const t = ctx.toks[ctx.pos];
  if (!t) throw new Error('Unexpected end of $filter input.');

  // Grouped sub-expression
  if (t.kind === 'lparen') {
    ctx.pos++;
    const inner = parseOrExpr(ctx);
    expect(ctx, 'rparen');
    return inner;
  }

  // Identifier — could be a Dataverse fn call, a string fn, a path leading to
  // a lambda, or the LHS of a comparison.
  if (t.kind === 'ident') {
    const ident = t.value;

    // ── Dataverse fn: Microsoft.Dynamics.CRM.X( ... )
    if (ident.startsWith('Microsoft.Dynamics.CRM.')) {
      // Must be followed by `(`.
      const next = ctx.toks[ctx.pos + 1];
      if (next?.kind === 'lparen') {
        return parseDvFunction(ctx, ident);
      }
    }

    // ── OData string fn: contains/startswith/endswith ( col, lit )
    if (STRING_FN_NAMES.has(ident)) {
      const next = ctx.toks[ctx.pos + 1];
      if (next?.kind === 'lparen') {
        return parseStringFunction(ctx, ident);
      }
    }

    // ── Lambda: <path>/any( [alias: inner] ) | <path>/all( ... )
    // The path may end in '/any' or '/all' — split and check.
    const lambdaMatch = ident.match(/^(.+?)\/(any|all)$/);
    if (lambdaMatch) {
      const next = ctx.toks[ctx.pos + 1];
      if (next?.kind === 'lparen') {
        return parseLambda(ctx, lambdaMatch[1], lambdaMatch[2] as 'any' | 'all');
      }
    }

    // ── Unknown function call — ident followed by `(` that didn't match
    //    any of the function shapes above. Typo'd Dataverse function name
    //    (missing namespace prefix), or made-up function. The user
    //    probably meant a real function but mis-spelled it. Throwing
    //    gives a clearer error than letting parseComparison choke on
    //    the unexpected `(` token.
    const next = ctx.toks[ctx.pos + 1];
    if (next?.kind === 'lparen') {
      throw new ParseFailure(
        `Unknown function \`${ident}(…)\` in $filter. Dataverse query functions must be prefixed with \`Microsoft.Dynamics.CRM.\` (e.g. \`Microsoft.Dynamics.CRM.${ident}(…)\`); OData string functions are \`contains\`/\`startswith\`/\`endswith\` only.`,
        DOCS.queryFunctions,
      );
    }

    // ── Comparison (eq/ne/…) — LHS is the ident; consume and continue.
    return parseComparison(ctx);
  }

  // String / number / bracket at this position is unexpected — treat as
  // best-effort recovery: produce a stub rule.
  throw new Error(`Unexpected token \`${t.value}\` at position ${ctx.pos} in $filter.`);
}

function parseComparison(ctx: ParseFilterCtx): FilterRule {
  const lhsTok = ctx.toks[ctx.pos++];
  if (lhsTok.kind !== 'ident') {
    throw new Error(`Comparison LHS must be a column reference, got \`${lhsTok.value}\`.`);
  }
  let col = lhsTok.value;
  // Decode wire-form lookup column to bare logical name. Preserve any
  // lambda-alias prefix or nav-path prefix.
  col = rewriteLookupValueInPath(col);

  const opTok = ctx.toks[ctx.pos++];
  if (!opTok || opTok.kind !== 'ident' || !COMPARISON_OPS.has(opTok.value.toLowerCase())) {
    throw new Error(`Expected comparison operator after \`${col}\`, got \`${opTok?.value}\`.`);
  }
  const op = opTok.value.toLowerCase();

  // Check for `eq null` / `ne null`
  const peek = ctx.toks[ctx.pos];
  if (peek?.kind === 'ident' && peek.value.toLowerCase() === 'null') {
    ctx.pos++;
    const nullOp = op === 'eq' ? 'is-null' : op === 'ne' ? 'is-not-null' : null;
    if (!nullOp) {
      ctx.warnings.push(iss(
        `\`${op} null\` is unusual — DRS will store as a literal value comparison.`,
        DOCS.filterRows,
      ));
    } else {
      return { id: newId('r'), type: 'rule', col, op: nullOp, val: '' };
    }
  }

  const rhs = ctx.toks[ctx.pos++];
  if (!rhs) throw new Error(`Comparison \`${col} ${op}\` is missing its RHS.`);

  // RHS classification:
  //   string literal       → valKind: literal, val = string
  //   number               → valKind: literal, val = number-as-string
  //   ident (true/false)   → valKind: literal, val = 'true'|'false'
  //   ident (path)         → valKind: column (same-table col-vs-col)
  //   ident (datetime)     → valKind: literal, val = ident
  //   ident (guid)         → valKind: literal, val = ident
  if (rhs.kind === 'string') {
    return { id: newId('r'), type: 'rule', col, op, val: rhs.value, valKind: 'literal' };
  }
  if (rhs.kind === 'number') {
    return { id: newId('r'), type: 'rule', col, op, val: rhs.value, valKind: 'literal' };
  }
  if (rhs.kind === 'ident') {
    const v = rhs.value;
    if (BOOL_LITERALS.has(v.toLowerCase())) {
      return { id: newId('r'), type: 'rule', col, op, val: v.toLowerCase(), valKind: 'literal' };
    }
    // GUID / date / numeric-looking ident → literal.
    if (looksLikeGuid(v) || looksLikeIsoDate(v)) {
      return { id: newId('r'), type: 'rule', col, op, val: v, valKind: 'literal' };
    }
    // Otherwise: column-vs-column comparison. Decode wire form on RHS too.
    return { id: newId('r'), type: 'rule', col, op, val: rewriteLookupValueInPath(v), valKind: 'column' };
  }
  throw new Error(`Unexpected RHS token \`${rhs.value}\` for \`${col} ${op}\`.`);
}

function parseStringFunction(ctx: ParseFilterCtx, name: string): FilterRule {
  // Already at ident; consume ident + '(' + col + ',' + lit + ')'.
  ctx.pos++; // consume ident
  expect(ctx, 'lparen');
  const colTok = ctx.toks[ctx.pos++];
  if (colTok.kind !== 'ident') throw new Error(`${name}( … ) expected a column reference first.`);
  expect(ctx, 'comma');
  const litTok = ctx.toks[ctx.pos++];
  if (litTok.kind !== 'string') throw new Error(`${name}(${colTok.value}, '…') expected a string literal as the second argument.`);
  expect(ctx, 'rparen');
  return {
    id: newId('r'),
    type: 'rule',
    col: rewriteLookupValueInPath(colTok.value),
    op: name,
    val: litTok.value,
    valKind: 'literal',
  };
}

function parseDvFunction(ctx: ParseFilterCtx, fullName: string): FilterFunctionNode {
  // fullName is like 'Microsoft.Dynamics.CRM.LastXDays'. Strip prefix.
  const shortName = fullName.replace(/^Microsoft\.Dynamics\.CRM\./, '');
  const opDef = OPERATORS.find(o => o.id === shortName);
  ctx.pos++; // consume the ident
  expect(ctx, 'lparen');

  // Parse named arguments — `PropertyName='col', PropertyValue=…, PropertyValues=[...]`.
  const args = parseDvFnArgs(ctx);
  expect(ctx, 'rparen');

  if (!opDef) {
    // Hard reject: an unknown `Microsoft.Dynamics.CRM.<X>` function is
    // either a typo or something not modeled in DRS. Dataverse will
    // reject too — applying it to the builder would just be a broken
    // chip the user can't fix.
    ctx.errors.push(iss(
      `Unknown Dataverse query function \`Microsoft.Dynamics.CRM.${shortName}\` — not in the documented set.`,
      DOCS.queryFunctions,
    ));
  }

  // Build the FilterFunctionNode. The encoder will strip alias prefixes / map
  // hierarchy fn to PK — we store what the user wrote.
  const node: FilterFunctionNode = {
    id: newId('f'),
    type: 'function',
    op: shortName,
    col: rewriteLookupValueInPath(args.PropertyName ?? ''),
  };
  if (args.PropertyValue != null) node.val = args.PropertyValue;
  if (args.PropertyValue1 != null && args.PropertyValue2 != null) {
    node.vals = [args.PropertyValue1, args.PropertyValue2];
  }
  if (args.PropertyValues) node.values = args.PropertyValues;
  return node;
}

interface DvFnArgs {
  PropertyName?: string;
  PropertyValue?: string;
  PropertyValue1?: string;
  PropertyValue2?: string;
  PropertyValues?: string[];
}

function parseDvFnArgs(ctx: ParseFilterCtx): DvFnArgs {
  const out: DvFnArgs = {};
  // Each arg: `Name=Value` or `Name='Value'` or `Name=[...]`. Separated by `,`.
  // Use the comma-after-(close-paren-of-depth-0) logic since values can be
  // bracket arrays which the tokenizer already groups.
  while (ctx.toks[ctx.pos] && ctx.toks[ctx.pos].kind !== 'rparen') {
    const nameTok = ctx.toks[ctx.pos++];
    if (nameTok.kind !== 'ident') throw new Error(`Expected argument name, got \`${nameTok.value}\`.`);
    // The argument name might come WITH `=Value` glued on as one token, since
    // the tokenizer doesn't split on `=`. Handle both shapes.
    let name = nameTok.value;
    let valueRaw: string | undefined;
    const eqIdx = name.indexOf('=');
    if (eqIdx >= 0) {
      valueRaw = name.slice(eqIdx + 1);
      name = name.slice(0, eqIdx);
    } else {
      // The `=` is in the next token's prefix — but the tokenizer should not
      // have created an `=` token. In practice the tokenizer groups
      // `PropertyName='createdon'` into one ident `PropertyName=` then a
      // string `createdon`. Let's check for that.
      // (We've trimmed identifiers on ',' / ':' / '(' / ')' / quote / bracket
      // so `PropertyName='x'` becomes ident=`PropertyName=`, string=`x`. Adjust.)
    }
    // Strip trailing `=` if present.
    if (name.endsWith('=')) name = name.slice(0, -1);

    let argValue: { kind: 'scalar'; v: string } | { kind: 'array'; v: string[] };
    if (valueRaw !== undefined && valueRaw !== '') {
      argValue = { kind: 'scalar', v: stripQuotes(valueRaw) };
    } else {
      const valTok = ctx.toks[ctx.pos++];
      if (!valTok) throw new Error(`Argument \`${name}\` is missing a value.`);
      if (valTok.kind === 'string')      argValue = { kind: 'scalar', v: valTok.value };
      else if (valTok.kind === 'number') argValue = { kind: 'scalar', v: valTok.value };
      else if (valTok.kind === 'ident')  argValue = { kind: 'scalar', v: valTok.value };
      else if (valTok.kind === 'bracket') argValue = { kind: 'array', v: parseBracketArray(valTok.value) };
      else throw new Error(`Unexpected token \`${valTok.value}\` for argument \`${name}\`.`);
    }

    if (name === 'PropertyName')        out.PropertyName        = argValue.kind === 'scalar' ? argValue.v : '';
    else if (name === 'PropertyValue')  out.PropertyValue       = argValue.kind === 'scalar' ? argValue.v : '';
    else if (name === 'PropertyValue1') out.PropertyValue1      = argValue.kind === 'scalar' ? argValue.v : '';
    else if (name === 'PropertyValue2') out.PropertyValue2      = argValue.kind === 'scalar' ? argValue.v : '';
    else if (name === 'PropertyValues') out.PropertyValues      = argValue.kind === 'array'  ? argValue.v : [];
    // Unknown arg → ignored silently (caller surfaces via the unknown-fn warning).

    // Consume trailing comma if present.
    if (ctx.toks[ctx.pos]?.kind === 'comma') ctx.pos++;
  }
  return out;
}

/** `['1000','5000']` → ['1000', '5000']. Tolerant of unquoted ints. */
function parseBracketArray(raw: string): string[] {
  try {
    // Dataverse PropertyValues uses single-quoted strings inside the array
    // (e.g. `['1000','5000']`). JSON requires double quotes. Convert.
    const jsonish = raw.replace(/'/g, '"');
    const parsed = JSON.parse(jsonish);
    return Array.isArray(parsed) ? parsed.map(v => String(v)) : [String(parsed)];
  } catch {
    // Fallback: strip brackets, split by comma, trim quotes.
    return raw
      .replace(/^\[/, '').replace(/\]$/, '')
      .split(',')
      .map(s => stripQuotes(s.trim()))
      .filter(Boolean);
  }
}

function stripQuotes(s: string): string {
  if (s.length >= 2 && s.startsWith("'") && s.endsWith("'")) return s.slice(1, -1);
  if (s.length >= 2 && s.startsWith('"') && s.endsWith('"')) return s.slice(1, -1);
  return s;
}

function parseLambda(ctx: ParseFilterCtx, navPath: string, lambdaKind: 'any' | 'all'): FilterLambdaNode {
  ctx.pos++; // consume the path ident
  expect(ctx, 'lparen');

  // navPath may carry an outer alias prefix (e.g. `c/Contact_Tasks`); strip it.
  let nav = navPath;
  if (ctx.outerAlias && nav.startsWith(ctx.outerAlias + '/')) {
    nav = nav.slice(ctx.outerAlias.length + 1);
  }

  // Empty lambda `any()` — no alias, no body.
  if (ctx.toks[ctx.pos]?.kind === 'rparen') {
    ctx.pos++;
    return {
      id: newId('l'),
      type: 'lambda',
      nav,
      lambda: lambdaKind,
      alias: nav[0] ?? 'x',
      inner: emptyFilterGroup(),
    };
  }

  // Otherwise: <alias> : <inner predicate>
  const aliasTok = ctx.toks[ctx.pos++];
  if (aliasTok.kind !== 'ident') throw new Error(`Lambda expected an alias, got \`${aliasTok.value}\`.`);
  const alias = aliasTok.value;
  expect(ctx, 'colon');

  // Parse inner predicate up to the matching ')'. Recurse with this alias
  // installed as outerAlias so nested lambdas can strip it from their nav.
  const savedOuter = ctx.outerAlias;
  ctx.outerAlias = alias;
  const inner = parseOrExpr(ctx);
  ctx.outerAlias = savedOuter;
  expect(ctx, 'rparen');

  const innerGroup: FilterGroup = inner.type === 'group'
    ? inner
    : { id: newId('g'), type: 'group', combinator: 'and', rules: [inner] };

  return {
    id: newId('l'),
    type: 'lambda',
    nav,
    lambda: lambdaKind,
    alias,
    inner: innerGroup,
  };
}

// ── Filter helpers ──────────────────────────────────────────────────────

function rewriteLookupValueInPath(p: string): string {
  // Walk segments; decode `_<x>_value` on each one.
  // Filter contexts don't easily know the leaf table when walking a
  // nav-path (would need recursive nav-walking metadata), so we run
  // without a table. Following the same conservative-preserve rule as
  // decodeLookupValueForm: when metadata is unavailable, leave the
  // `_X_value` wire form intact so the encoder's pass-through behavior
  // produces a faithful round-trip. Decoding without metadata risks
  // emitting a bare name the encoder can't re-wrap.
  return p.split('/').map(seg => decodeLookupValueForm(seg)).join('/');
}

function looksLikeGuid(s: string): boolean {
  return /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(s);
}

function looksLikeIsoDate(s: string): boolean {
  // Accepts:
  //   2024-01-01                            date-only
  //   2024-01-01T00:00                      datetime w/o seconds
  //   2024-01-01T00:00:00                   datetime w/ seconds
  //   2024-01-01T00:00:00.123                datetime w/ fractional seconds
  //   any of the above + 'Z' | '+05:30' | '-08:00'  zone designator
  return /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2}(\.\d+)?)?(Z|[+-]\d{2}:\d{2})?)?$/.test(s);
}

function isKeywordAhead(ctx: ParseFilterCtx, kw: string): boolean {
  const t = ctx.toks[ctx.pos];
  return !!t && t.kind === 'ident' && t.value.toLowerCase() === kw;
}

function consumeKeyword(ctx: ParseFilterCtx, kw: string): void {
  if (!isKeywordAhead(ctx, kw)) throw new Error(`Expected \`${kw}\` at position ${ctx.pos}.`);
  ctx.pos++;
}

function expect(ctx: ParseFilterCtx, kind: Tok['kind']): void {
  const t = ctx.toks[ctx.pos];
  if (!t || t.kind !== kind) throw new Error(`Expected ${kind} at position ${ctx.pos}, got ${t?.kind ?? 'EOF'}.`);
  ctx.pos++;
}

// ── $expand parser ──────────────────────────────────────────────────────

function parseExpand(input: string, parentTable: TableMeta | undefined): { items: ExpandSpec[]; warnings: ParseIssue[]; errors: ParseIssue[] } {
  const warnings: ParseIssue[] = [];
  const errors: ParseIssue[] = [];
  // Top-level comma-separated expand items.
  const parts = topLevelSplit(input, ',').map(s => s.trim()).filter(Boolean);
  const items = parts.map(p => parseOneExpand(p, parentTable, warnings, errors)).filter((x): x is ExpandSpec => !!x);
  return { items, warnings, errors };
}

function parseOneExpand(raw: string, parentTable: TableMeta | undefined, warnings: ParseIssue[], errors: ParseIssue[]): ExpandSpec | null {
  // `<nav>` or `<nav>(opts)` or `<nav>/$ref`.
  const navOpenIdx = raw.indexOf('(');
  let nav: string;
  let optsRaw = '';
  if (navOpenIdx < 0) {
    nav = raw.trim();
  } else {
    nav = raw.slice(0, navOpenIdx).trim();
    // Strip surrounding parens around opts.
    optsRaw = raw.slice(navOpenIdx + 1, raw.lastIndexOf(')')).trim();
  }

  // `/$ref` suffix on a single-valued nav → we don't model $ref in DRS yet.
  if (nav.endsWith('/$ref')) {
    nav = nav.slice(0, -'/$ref'.length);
    warnings.push(iss(
      `\`$expand=${nav}/$ref\` was reduced to a plain \`$expand=${nav}\` — DRS doesn't render the $ref-only form.`,
      DOCS.joinTables,
    ));
  }

  const spec = makeExpandSpec(nav);

  if (!optsRaw) return spec;

  const navMeta = parentTable?.navigationProperties.find(n => n.name === nav);
  const targetTbl = navMeta ? findTable(navMeta.targetEntity) : undefined;

  // Inner options are `$key=value` segments separated by `;`.
  const innerOpts = topLevelSplit(optsRaw, ';').map(s => s.trim()).filter(Boolean);
  for (const opt of innerOpts) {
    const eq = opt.indexOf('=');
    if (eq < 0) continue;
    const k = opt.slice(0, eq);
    const v = opt.slice(eq + 1);
    switch (k) {
      case '$select':  spec.select = parseSelect(v, targetTbl); break;
      case '$orderby': spec.orderby = parseOrderby(v, targetTbl); break;
      case '$top':     spec.top = parseIntOrNull(v); break;
      case '$filter': {
        const r = parseFilter(v, targetTbl);
        spec.filter = r.group;
        warnings.push(...r.warnings.map(w => prefixIss(`expand(${nav})`, w)));
        errors.push(...r.errors.map(e => prefixIss(`expand(${nav})`, e)));
        break;
      }
      case '$expand': {
        const r = parseExpand(v, targetTbl);
        spec.nestedExpand = r.items;
        warnings.push(...r.warnings);
        errors.push(...r.errors);
        break;
      }
      default:
        errors.push(iss(
          `expand(${nav}): unsupported inner option \`${k}\` — DRS supports $select / $filter / $orderby / $top / $expand only.`,
          DOCS.joinTables,
        ));
    }
  }
  return spec;
}

// ── $apply parser ───────────────────────────────────────────────────────

function parseApply(input: string, table: TableMeta | undefined): { spec: ApplySpec; warnings: ParseIssue[]; errors: ParseIssue[] } {
  const warnings: ParseIssue[] = [];
  const errors: ParseIssue[] = [];
  const spec: ApplySpec = {
    enabled: true,
    prefilter: emptyFilterGroup(),
    groupby: [],
    aggregates: [],
  };
  const stages = topLevelSplit(input, '/').map(s => s.trim()).filter(Boolean);
  for (const stage of stages) {
    if (stage.startsWith('filter(') && stage.endsWith(')')) {
      const body = stage.slice('filter('.length, -1);
      const r = parseFilter(body, table);
      spec.prefilter = r.group;
      warnings.push(...r.warnings.map(w => prefixIss('apply/filter', w)));
      errors.push(...r.errors.map(e => prefixIss('apply/filter', e)));
    } else if (stage.startsWith('aggregate(') && stage.endsWith(')')) {
      const body = stage.slice('aggregate('.length, -1);
      spec.aggregates = parseAggregateList(body, warnings, errors);
    } else if (stage.startsWith('groupby(') && stage.endsWith(')')) {
      const body = stage.slice('groupby('.length, -1);
      const r = parseGroupby(body, warnings, errors);
      spec.groupby = r.cols;
      if (r.aggregates) spec.aggregates = r.aggregates;
    } else {
      // Unknown $apply stage — `bottomcount`, `topcount`, `compute`, etc.
      // Hard reject — DRS only models filter/groupby/aggregate.
      errors.push(iss(
        `$apply: unsupported stage \`${stage.split('(')[0]}(…)\`.`,
        DOCS.aggregateData,
      ));
    }
  }
  return { spec, warnings, errors };
}

function parseGroupby(body: string, warnings: ParseIssue[], errors: ParseIssue[]): { cols: string[]; aggregates?: ApplyAgg[] } {
  const firstOpen = body.indexOf('(');
  if (firstOpen !== 0) {
    errors.push(iss(
      `$apply/groupby: expected \`(\` at start of args, got \`${body[0] ?? 'EOF'}\`.`,
      DOCS.aggregateData,
    ));
    return { cols: [] };
  }
  let depth = 0; let firstCloseIdx = -1;
  for (let i = 0; i < body.length; i++) {
    if (body[i] === '(') depth++;
    else if (body[i] === ')') { depth--; if (depth === 0) { firstCloseIdx = i; break; } }
  }
  if (firstCloseIdx < 0) {
    errors.push(iss('$apply/groupby: unbalanced parens.', DOCS.aggregateData));
    return { cols: [] };
  }
  const colsRaw = body.slice(1, firstCloseIdx);
  const cols = topLevelSplit(colsRaw, ',').map(s => s.trim()).filter(Boolean)
    .map(c => rewriteLookupValueInPath(c));

  const rest = body.slice(firstCloseIdx + 1).trim();
  if (!rest) return { cols };
  if (rest.startsWith(',aggregate(') && rest.endsWith(')')) {
    const aggBody = rest.slice(',aggregate('.length, -1);
    return { cols, aggregates: parseAggregateList(aggBody, warnings, errors) };
  }
  errors.push(iss(
    `$apply/groupby: unexpected trailing content after the cols list.`,
    DOCS.aggregateData,
  ));
  return { cols };
}

function parseAggregateList(body: string, _warnings: ParseIssue[], errors: ParseIssue[]): ApplyAgg[] {
  const parts = topLevelSplit(body, ',').map(s => s.trim()).filter(Boolean);
  const out: ApplyAgg[] = [];
  for (const p of parts) {
    const countMatch = p.match(/^\$count\s+as\s+(\S+)$/i);
    if (countMatch) {
      out.push({ col: '$count', fn: '$count', alias: countMatch[1] });
      continue;
    }
    const m = p.match(/^(\S+)\s+with\s+(\w+)\s+as\s+(\S+)$/i);
    if (!m) {
      errors.push(iss(
        `$apply/aggregate: couldn't parse \`${p}\` — expected \`<col> with <fn> as <alias>\` or \`$count as <alias>\`.`,
        DOCS.aggregateData,
      ));
      continue;
    }
    const fn = m[2].toLowerCase();
    if (!KNOWN_AGG_FNS.has(fn)) {
      errors.push(iss(
        `$apply/aggregate: function \`${fn}\` is not modeled by DRS — only sum / average / min / max / $count are supported.`,
        DOCS.aggregateData,
      ));
      continue;
    }
    out.push({
      col: rewriteLookupValueInPath(m[1]),
      fn: fn as AggFn,
      alias: m[3],
    });
  }
  return out;
}

// ── Numeric helpers ─────────────────────────────────────────────────────

function parseIntOrNull(s: string): number | null {
  const n = parseInt(s, 10);
  return Number.isFinite(n) ? n : null;
}
