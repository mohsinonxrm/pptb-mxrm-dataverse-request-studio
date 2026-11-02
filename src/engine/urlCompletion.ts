// Completion for hand-typed Web API URLs (#268): the Batch custom URL and
// query options, the Paste OData URL dialog, and Retrieve Next Link.
//
// Two steps, both pure:
//   1. `urlCaretContext` reads the text up to the caret and says what's
//      being typed there (a table, a navigation property, a query option,
//      a column in $select, …), for which table, and the range to replace.
//      Syntax only: it never looks at metadata.
//   2. `urlSuggestions` turns that into ranked suggestions from the
//      metadata, or names the table it still needs loaded.
//
// Unlike `parseODataUrl`, which validates a whole URL, this has to cope
// with the half-typed text in front of the caret.

import type { ColumnMeta, NavProperty, TableMeta } from '../metadata/metadata';
import type { CsdlAction } from '../metadata/actionsCsdl';

/** `url`: a whole URL or path. `query`: query options only, for a known table. */
export type UrlMode = 'url' | 'query';

/** A table reached from an entity set through navigation properties. */
export interface TableRef {
  entitySet: string;
  navs: string[];
}

/** What a position can take. */
export type Want =
  | 'tables' // entity sets and unbound operations
  | 'segment' // after a table: navs and bound operations
  | 'options'
  | 'nestedOptions'
  | 'columns'
  | 'navs'
  | 'operand' // in $filter: columns, single-valued navs, functions
  | 'operators'
  | 'logical'
  | 'orderDir'
  | 'booleans';

export interface CaretContext {
  want: Want;
  /** The text the suggestion replaces: [from, to). */
  from: number;
  to: number;
  /** What's typed so far in that range, before the caret. */
  prefix: string;
  /** The table the suggestions come from (unset for tables, operations and options at the root). */
  table?: TableRef;
  /** For `segment`: the path addresses one row (a key, or a single-valued nav). */
  keyed?: boolean;
  /** For `operand`: how many of `table.navs` the `$filter` path added (`nav/col`). */
  pathNavs?: number;
}

const IDENT_TAIL = /^[A-Za-z0-9_$.@]*/;

interface Scan {
  /** Unclosed `(` positions, outermost first. */
  open: number[];
  inQuote: boolean;
  /** The last position of one of `seps` outside parentheses and quotes, or -1. */
  lastSep: (seps: string) => number;
}

/** Walk `text[start, end)`, tracking parentheses and `'…'` strings. */
function scan(text: string, start: number, end: number): Scan {
  const open: number[] = [];
  const seps: { i: number; ch: string }[] = [];
  let inQuote = false;
  for (let i = start; i < end; i++) {
    const ch = text[i];
    if (inQuote) {
      // '' is an escaped quote; at the end of the text it may be a closing
      // quote the caret sits after, which reads the same
      if (ch === "'") inQuote = false;
      continue;
    }
    if (ch === "'") inQuote = true;
    else if (ch === '(') open.push(i);
    else if (ch === ')') open.pop();
    else if (open.length === 0) seps.push({ i, ch });
  }
  return {
    open,
    inQuote,
    lastSep: (chars) => {
      for (let k = seps.length - 1; k >= 0; k--) if (chars.includes(seps[k].ch)) return seps[k].i;
      return -1;
    },
  };
}

/** Name before a segment's key, e.g. `accounts` in `accounts(…)`. */
const segName = (seg: string) => seg.replace(/\(.*$/s, '').trim();

/**
 * Where the path starts, after `https://host/api/data/v9.2/`,
 * `/api/data/v9.2/` or a leading `/`. -1 while the host or the API base is
 * still being typed: nothing is offered there.
 */
function pathStart(text: string): number {
  let i = text.match(/^\s*/)![0].length;
  const host = text.slice(i).match(/^https?:\/\/[^/?#\s]*\//i);
  if (host) i += host[0].length;
  else if (/^https?(?::|$)/i.test(text.slice(i))) return -1;
  else if (text[i] === '/') i++;
  const rest = text.slice(i);
  const base = rest.match(/^api\/data\/v\d+(?:\.\d+)?\//i);
  if (base) return i + base[0].length;
  // An absolute URL needs its API base; `api/…` is one on its way
  if (host || /^api\/(?:d(?:a(?:t(?:a(?:\/(?:v[\d.]*)?)?)?)?)?)?$/i.test(rest)) return -1;
  return i;
}

/**
 * What's being typed at `caret` in `text`. `entitySet` is the table for
 * `query` mode. Null where nothing is offered (a key, a string, a value).
 */
export function urlCaretContext(
  text: string,
  caret: number,
  mode: UrlMode,
  entitySet?: string,
): CaretContext | null {
  const to = caret + (text.slice(caret).match(IDENT_TAIL)?.[0].length ?? 0);
  if (mode === 'query') {
    if (!entitySet?.trim()) return optionsContext(text, 0, caret, to, '&', null);
    return optionsContext(text, 0, caret, to, '&', { entitySet: entitySet.trim(), navs: [] });
  }

  const start = pathStart(text);
  if (start < 0 || caret < start) return null;
  // Path segments, up to the `?` (outside keys and strings)
  const segStarts = [start];
  let depth = 0;
  let inQuote = false;
  for (let i = start; i < caret; i++) {
    const ch = text[i];
    if (inQuote) {
      if (ch === "'") inQuote = false;
      continue;
    }
    if (ch === "'") inQuote = true;
    else if (ch === '(') depth++;
    else if (ch === ')') depth--;
    else if (depth === 0 && ch === '/') segStarts.push(i + 1);
    else if (depth === 0 && ch === '?') {
      const table = tableOfPath(text, segStarts, i);
      return optionsContext(text, i + 1, caret, to, '&', table);
    }
  }
  if (depth > 0 || inQuote) return null; // in a key
  const segStart = segStarts[segStarts.length - 1];
  const prefix = text.slice(segStart, caret);
  if (segStarts.length === 1) {
    // $1 (a Content-ID), $batch, $metadata
    if (prefix.startsWith('$')) return null;
    return { want: 'tables', from: segStart, to, prefix };
  }
  const prev = text.slice(segStarts[segStarts.length - 2], segStart - 1);
  const table = tableOfPath(text, segStarts.slice(0, -1), segStart - 1);
  if (!table) return null;
  return { want: 'segment', from: segStart, to, prefix, table, keyed: prev.includes('(') };
}

/** The table a path addresses, or null past an operation, `$ref` or `$count`. */
function tableOfPath(text: string, segStarts: number[], end: number): TableRef | null {
  const segs = segStarts.map((s, i) =>
    segName(text.slice(s, i + 1 < segStarts.length ? segStarts[i + 1] - 1 : end)),
  );
  if (!segs[0] || segs.some((s) => !s || s.startsWith('$') || s.includes('.'))) return null;
  return { entitySet: segs[0], navs: segs.slice(1) };
}

/** Inside a query string (`sep` `&`) or an `$expand(…)` (`sep` `;`). */
function optionsContext(
  text: string,
  start: number,
  caret: number,
  to: number,
  sep: '&' | ';',
  table: TableRef | null,
): CaretContext | null {
  const s = scan(text, start, caret);
  if (s.inQuote) return null;
  const optStart = Math.max(start, s.lastSep(sep) + 1);
  const opt = text.slice(optStart, caret);
  if (!opt.includes('=')) {
    if (s.open.length) return null;
    return {
      want: sep === ';' ? 'nestedOptions' : 'options',
      from: optStart + (opt.length - opt.trimStart().length),
      to,
      prefix: opt.trimStart(),
      ...(table ? { table } : {}),
    };
  }
  return valueContext(text, optStart, caret, to, table);
}

/** In an option's value: `$select=…`, `$expand=…`, … */
function valueContext(
  text: string,
  optStart: number,
  caret: number,
  to: number,
  table: TableRef | null,
): CaretContext | null {
  const eq = text.indexOf('=', optStart);
  if (eq < 0 || eq >= caret || !table) return null;
  const name = text.slice(optStart, eq).trim().toLowerCase();
  const valueStart = eq + 1;
  const s = scan(text, valueStart, caret);
  if (s.inQuote) return null;
  const itemStart = Math.max(valueStart, s.lastSep(',') + 1);

  if (name === '$expand') {
    if (s.open.length) {
      // In `nav(…)`: that nav's own options, `;`-separated
      const nav = text.slice(itemStart, s.open[0]).trim();
      if (!nav) return null;
      return optionsContext(text, s.open[0] + 1, caret, to, ';', {
        ...table,
        navs: [...table.navs, nav],
      });
    }
    const prefix = text.slice(itemStart, caret).trimStart();
    return { want: 'navs', from: caret - prefix.length, to, prefix, table };
  }

  if (name === '$select' || name === '$orderby') {
    if (s.open.length) return null;
    const item = text.slice(itemStart, caret).trimStart();
    if (name === '$orderby' && /\s/.test(item)) {
      // `name desc `: the item is complete
      if (/^\S+\s+\S+\s+$/.test(item)) return null;
      const dir = item.slice(item.search(/\S*$/));
      return { want: 'orderDir', from: caret - dir.length, to, prefix: dir, table };
    }
    return { want: 'columns', from: caret - item.length, to, prefix: item, table };
  }

  if (name === '$filter') return filterContext(text, valueStart, caret, to, table);
  if (name === '$count') {
    return { want: 'booleans', from: valueStart, to, prefix: text.slice(valueStart, caret), table };
  }
  return null;
}

const FILTER_KEYWORDS = new Set([
  'and',
  'or',
  'not',
  'eq',
  'ne',
  'gt',
  'ge',
  'lt',
  'le',
  'null',
  'true',
  'false',
]);

function filterContext(
  text: string,
  valueStart: number,
  caret: number,
  to: number,
  table: TableRef,
): CaretContext | null {
  const s = scan(text, valueStart, caret);
  if (s.inQuote) return null;
  const head = text.slice(valueStart, caret);
  const word = head.match(/[A-Za-z0-9_/]*$/)?.[0] ?? '';
  const before = head.slice(0, head.length - word.length);
  // `-` joins GUIDs and negative numbers into one token
  const prevTok = before.trimEnd().match(/(?:[A-Za-z0-9_/.-]+|'(?:[^']|'')*'|.)$/)?.[0] ?? '';
  // After a comma comes a function's value argument (`contains(name, 'x')`)
  if (prevTok === ',') return null;
  const operandPos =
    prevTok === '' || prevTok === '(' || ['and', 'or', 'not'].includes(prevTok.toLowerCase());
  if (operandPos) {
    // `nav/col`: the column of a single-valued nav's table
    const slash = word.lastIndexOf('/');
    if (slash >= 0) {
      const navs = word.slice(0, slash).split('/');
      const prefix = word.slice(slash + 1);
      return {
        want: 'operand',
        from: caret - prefix.length,
        to,
        prefix,
        table: { ...table, navs: [...table.navs, ...navs] },
        pathNavs: navs.length,
      };
    }
    return { want: 'operand', from: caret - word.length, to, prefix: word, table };
  }
  // After a space: an operator after a column, `and`/`or` after a value
  if (before.length === 0 || !/\s$/.test(before)) return null;
  const isValue =
    prevTok === ')' ||
    prevTok.startsWith("'") ||
    /^-?\d/.test(prevTok) ||
    /^[0-9a-f]{8}-[0-9a-f]{4}-/i.test(prevTok) ||
    ['null', 'true', 'false'].includes(prevTok.toLowerCase());
  if (isValue) return { want: 'logical', from: caret - word.length, to, prefix: word, table };
  if (FILTER_KEYWORDS.has(prevTok.toLowerCase())) return null;
  return { want: 'operators', from: caret - word.length, to, prefix: word, table };
}

// ── Suggestions ────────────────────────────────────────────────────────

export interface UrlSuggestion {
  /** Shown in the list. */
  label: string;
  /** Replaces the context's range. */
  insert: string;
  /** Secondary text: display name, type, target. */
  detail?: string;
  /** Also matched when typing: a table's or column's display name. */
  displayName?: string;
  kind: 'table' | 'operation' | 'nav' | 'column' | 'option' | 'keyword';
}

type EntityItem = { logicalName: string; entitySetName: string; displayName: string };

export interface SuggestionDeps {
  /** The tables listed (the pickers' scope). */
  entities: readonly EntityItem[];
  /** Every table, for reading a typed entity set. Default: `entities`. */
  allEntities?: readonly EntityItem[];
  /** A loaded table, or undefined while it isn't. */
  tableOf: (logicalName: string) => TableMeta | undefined;
  /** The CSDL's functions and actions; empty until loaded. */
  operations: readonly CsdlAction[];
  columnInScope?: (c: ColumnMeta) => boolean;
  navInScope?: (n: NavProperty) => boolean;
}

export interface SuggestionResult {
  items: UrlSuggestion[];
  /** More matches than `items` holds. */
  hiddenCount: number;
  /** A table to load before the suggestions are complete. */
  needTable?: string;
}

/** Most suggestions listed at once. */
export const MAX_SUGGESTIONS = 50;

const OPTIONS = ['$select', '$filter', '$orderby', '$expand', '$top', '$count', '$apply'];
const NESTED_OPTIONS = ['$select', '$filter', '$orderby', '$top', '$expand'];
const OPERATORS = ['eq', 'ne', 'gt', 'ge', 'lt', 'le'];
const FILTER_FUNCTIONS = ['contains', 'startswith', 'endswith'];

function resolveTable(
  ref: TableRef,
  deps: SuggestionDeps,
): { table?: TableMeta; path?: NavProperty[]; need?: string } {
  const root = (deps.allEntities ?? deps.entities).find(
    (e) => e.entitySetName.toLowerCase() === ref.entitySet.toLowerCase(),
  );
  if (!root) return {};
  let table = deps.tableOf(root.logicalName);
  if (!table) return { need: root.logicalName };
  const path: NavProperty[] = [];
  for (const name of ref.navs) {
    const nav = table.navigationProperties.find((n) => n.name === name);
    if (!nav) return {};
    path.push(nav);
    const next = deps.tableOf(nav.targetEntity);
    if (!next) return { need: nav.targetEntity };
    table = next;
  }
  return { table, path };
}

/**
 * A `$filter` query function (Last7Days, Contains, …), not a path segment.
 * Dataverse declares them unbound, so `isQueryFunction` (which looks for
 * IsBound) can miss them: match their signature too.
 */
const isFilterFunction = (a: CsdlAction) =>
  a.isQueryFunction ||
  (a.kind === 'Function' &&
    a.parameters[0]?.name === 'PropertyName' &&
    a.returnType?.typeName === 'Edm.Boolean');

const operationInsert = (a: CsdlAction, bound: boolean) => {
  const name = bound ? `${a.namespace}.${a.name}` : a.name;
  if (a.kind === 'Action') return name;
  // The CSDL provider leaves out a bound function's binding parameter
  return a.parameters.length ? `${name}(` : `${name}()`;
};

const columnDetail = (c: ColumnMeta) => `${c.displayName} (${c.attributeType})`;

/** Suggestions for a caret context, best match first. */
export function urlSuggestions(ctx: CaretContext, deps: SuggestionDeps): SuggestionResult {
  const colOk = deps.columnInScope ?? (() => true);
  const navOk = deps.navInScope ?? (() => true);
  const all: UrlSuggestion[] = [];
  let needTable: string | undefined;
  const resolved = ctx.table ? resolveTable(ctx.table, deps) : {};
  if (resolved.need) needTable = resolved.need;
  const table = resolved.table;
  const columns = (filter: (c: ColumnMeta) => boolean = () => true) =>
    (table?.columns ?? [])
      .filter((c) => colOk(c) && c.isValidForRead !== false && filter(c))
      .map(
        (c): UrlSuggestion => ({
          label: c.oDataName ?? c.logicalName,
          insert: c.oDataName ?? c.logicalName,
          detail: columnDetail(c),
          displayName: c.displayName,
          kind: 'column',
        }),
      )
      .sort((a, b) => bare(a.label).localeCompare(bare(b.label)));
  const navs = (filter: (n: NavProperty) => boolean = () => true) =>
    (table?.navigationProperties ?? [])
      .filter((n) => navOk(n) && filter(n))
      .map(
        (n): UrlSuggestion => ({
          label: n.name,
          insert: n.name,
          detail: `→ ${n.targetEntity}${n.cardinality === 'ManyToOne' ? '' : ' (collection)'}`,
          kind: 'nav',
        }),
      )
      .sort((a, b) => a.label.localeCompare(b.label));
  const keywords = (words: string[], detail?: string) =>
    words.map((w): UrlSuggestion => ({ label: w, insert: w, detail, kind: 'keyword' }));

  switch (ctx.want) {
    case 'tables':
      for (const e of [...deps.entities].sort((a, b) =>
        a.entitySetName.localeCompare(b.entitySetName),
      )) {
        all.push({
          label: e.entitySetName,
          insert: e.entitySetName,
          detail: e.displayName,
          displayName: e.displayName,
          kind: 'table',
        });
      }
      // An unbound function or action can start the path too
      for (const a of deps.operations) {
        if (a.binding.kind !== 'unbound' || isFilterFunction(a)) continue;
        all.push({
          label: a.name,
          insert: operationInsert(a, false),
          detail: a.kind,
          kind: 'operation',
        });
      }
      break;
    case 'segment': {
      if (!table) break;
      const lastNav = resolved.path?.[resolved.path.length - 1];
      const single = ctx.keyed || lastNav?.cardinality === 'ManyToOne';
      if (single) all.push(...navs());
      else all.push(...keywords(['$count'], 'Number of rows'));
      for (const a of deps.operations) {
        if (a.binding.kind === 'unbound' || isFilterFunction(a)) continue;
        // `crmbaseentity`: bound to any table
        const t = a.binding.entityType;
        if (t !== table.logicalName && t !== 'crmbaseentity') continue;
        if ((a.binding.kind === 'entity') !== !!single) continue;
        all.push({
          label: `${a.namespace}.${a.name}`,
          insert: operationInsert(a, true),
          detail: a.kind,
          kind: 'operation',
        });
      }
      break;
    }
    case 'options':
    case 'nestedOptions':
      for (const o of ctx.want === 'options' ? OPTIONS : NESTED_OPTIONS) {
        all.push({ label: o, insert: `${o}=`, kind: 'option' });
      }
      break;
    case 'columns':
      all.push(...columns());
      break;
    case 'navs':
      all.push(...navs());
      break;
    case 'operand':
      // `nav/col` needs single-valued navs; a collection takes any / all
      if (
        resolved.path
          ?.slice(resolved.path.length - (ctx.pathNavs ?? 0))
          .some((n) => n.cardinality !== 'ManyToOne')
      )
        break;
      all.push(...columns());
      all.push(
        ...navs((n) => n.cardinality === 'ManyToOne').map((s) => ({
          ...s,
          insert: `${s.insert}/`,
        })),
      );
      if (ctx.prefix.indexOf('/') < 0) {
        all.push(
          ...FILTER_FUNCTIONS.map(
            (f): UrlSuggestion => ({
              label: f,
              insert: `${f}(`,
              detail: 'Function',
              kind: 'keyword',
            }),
          ),
        );
      }
      break;
    case 'operators':
      all.push(...keywords(OPERATORS, 'Operator'));
      break;
    case 'logical':
      all.push(...keywords(['and', 'or']));
      break;
    case 'orderDir':
      all.push(...keywords(['asc', 'desc']));
      break;
    case 'booleans':
      all.push(...keywords(['true', 'false']));
      break;
  }

  return { ...rank(all, ctx.prefix), ...(needTable ? { needTable } : {}) };
}

/** `$select` → `select`, `_parentcustomerid_value` → `parentcustomerid_value`. */
const bare = (s: string) => s.replace(/^[$_]/, '');

/**
 * Labels starting with the prefix first, shortest first (`cont` →
 * `contacts` before `contactinvoicescollection`), then other matches;
 * capped. Nothing typed keeps the list's own order.
 */

function rank(
  all: UrlSuggestion[],
  prefix: string,
): { items: UrlSuggestion[]; hiddenCount: number } {
  const p = bare(prefix.trim().toLowerCase());
  const starts: UrlSuggestion[] = [];
  const contains: UrlSuggestion[] = [];
  for (const s of all) {
    const label = bare(s.label.toLowerCase());
    if (!p || label.startsWith(p)) starts.push(s);
    else if (label.includes(p) || s.displayName?.toLowerCase().includes(p)) contains.push(s);
  }
  // Stable: equal lengths keep their order
  if (p) starts.sort((a, b) => a.label.length - b.label.length);
  const matched = [...starts, ...contains];
  // A list holding only what's already typed has nothing to offer
  if (matched.length === 1 && matched[0].insert.toLowerCase() === prefix.trim().toLowerCase()) {
    return { items: [], hiddenCount: 0 };
  }
  return {
    items: matched.slice(0, MAX_SUGGESTIONS),
    hiddenCount: Math.max(0, matched.length - MAX_SUGGESTIONS),
  };
}

/**
 * The text with `s` inserted for `ctx`, and the caret after it. Typed
 * mid-word, the replaced range ends where the word does, so what already
 * follows it (`=`, `/`, `()`, `(`) isn't inserted a second time.
 */
export function applySuggestion(
  text: string,
  ctx: CaretContext,
  s: UrlSuggestion,
): { value: string; caret: number } {
  let insert = s.insert;
  let skip = 0;
  const after = text.slice(ctx.to);
  for (const tail of ['()', '(', '=', '/']) {
    if (insert.endsWith(tail) && after.startsWith(tail)) {
      insert = insert.slice(0, -tail.length);
      skip = tail.length;
      break;
    }
  }
  return {
    value: text.slice(0, ctx.from) + insert + after,
    caret: ctx.from + insert.length + skip,
  };
}
