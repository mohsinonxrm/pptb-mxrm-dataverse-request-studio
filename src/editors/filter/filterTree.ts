// Filter tree model + OData encoder.
//
// Four node kinds — matches the v2.2 visual vocabulary 1:1:
//   - 'group'    : combinator (and/or) + ordered children (any kind)
//   - 'rule'     : column + operator + value(s) — basic comparison/string/null-check
//   - 'function' : Dataverse query function — Microsoft.Dynamics.CRM.* with PropertyName/PropertyValue(s)
//   - 'lambda'   : collection-nav/any|all(alias: <nested predicate>) — over a 1:N or N:N navigation
//
// Operators in 'rule' nodes are pulled from operators.ts (comparison + string + null-check).
// Operators in 'function' nodes are the same registry filtered to dv-fn-* kinds.
// Lambda nodes own their own nested FilterGroup.

import type { ColumnMeta, TableMeta } from '../../mock/metadata';
import { resolveNavPath } from '../../mock/metadata';
import { findOperator, type OperatorDef } from './operators';
import { attrRef } from '../../engine/odataAttr';
// `Pick<>` here keeps the literal type narrow but allows ColumnMeta arity to widen
export type { ColumnMeta } from '../../mock/metadata';

export type Combinator = 'and' | 'or';

export interface FilterRule {
  id: string;
  type: 'rule';
  /** Column logical name (or `nav/col` / `alias/col` for traversed / lambda contexts) */
  col: string;
  /** Operator id (must be a 'comparison' / 'string' / 'null-check' kind) */
  op: string;
  /** RHS mode — `literal` (default) or `column` (bare property name on RHS). */
  valKind?: 'literal' | 'column';
  /**
   * RHS value.
   *  - literal mode: parsed per column AttributeType
   *  - column   mode: logical name of the comparison column on the same row
   */
  val?: string;
  /**
   * Per-condition negation. Only meaningful when `op` is one of the three OData
   * string functions (contains / startswith / endswith) — Dataverse does NOT
   * support negating arbitrary comparison rules or whole groups. See
   * Filter-Builder-Scenarios §2A.
   */
  negated?: boolean;
}

export interface FilterFunctionNode {
  id: string;
  type: 'function';
  /** Operator id (must be a 'dv-fn-*' kind) */
  op: string;
  /** PropertyName — column the function targets */
  col: string;
  /** Single PropertyValue (arity 1) */
  val?: string;
  /** Two PropertyValues (arity 2 — fiscal period+year, etc.) */
  vals?: [string, string];
  /** Array PropertyValues (arity n — In/NotIn/Between/ContainValues) */
  values?: string[];
  /**
   * Per-function negation — emits `not Microsoft.Dynamics.CRM.X(...)`. Per
   * the MS Learn queryfunctions reference, this is the documented escape hatch
   * for functions without a negative twin (Today, Last7Days, EqualUserId, …).
   */
  negated?: boolean;
}

export interface FilterLambdaNode {
  id: string;
  type: 'lambda';
  /** Collection navigation property name (e.g. contact_customer_accounts) */
  nav: string;
  /** any | all */
  lambda: 'any' | 'all';
  /** Single-letter alias bound inside (default derived from nav) */
  alias: string;
  /** Nested predicate — when empty, emits `nav/any()` */
  inner: FilterGroup;
  /**
   * Emit `not <nav>/<any|all>(...)` when true. Empirically supported per
   * the G.8 test (`not contact_customer_accounts/any(...)` → 200 OK on
   * v9.2). Mirrors the group-level `negated` flag — the toggle lives on
   * the lambda chip's header in the UI.
   */
  negated?: boolean;
}

export interface FilterGroup {
  id: string;
  type: 'group';
  combinator: Combinator;
  rules: FilterNode[];
  /**
   * Emit `not (<combined body>)` around the group's expression when true.
   * Empirically supported across the G.1–G.7 test set:
   *   • `not (a or b)`, `not (a and b)`, `not (3+ rules)`, `not (mixed)`,
   *     `not (a and not (b or c))`, and `<positive> and not (group)` all
   *     return 200 OK with the expected row counts.
   *
   * Caveat: `not (<group containing a Microsoft.Dynamics.CRM.* fn>)` is
   * rejected with 405 (test G.9) even with parens. `validateRequest`
   * surfaces a warning in that case — the encoder still emits faithfully.
   */
  negated?: boolean;
}

export type FilterNode = FilterRule | FilterFunctionNode | FilterLambdaNode | FilterGroup;

/**
 * Generate a unique id. Uses `crypto.randomUUID` when available (every modern
 * browser + Node 19+) — falls back to a counter+random pattern. Crucially the
 * IDs are GLOBALLY UNIQUE so they never collide with anything the seed tree
 * hardcodes, which used to cause edits to one node to visually "leak" into
 * another node with the same id.
 */
let __fallbackCounter = 1;
export function newId(prefix: string): string {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return `${prefix}_${crypto.randomUUID()}`;
  }
  return `${prefix}_${Date.now().toString(36)}_${(__fallbackCounter++).toString(36)}_${Math.floor(Math.random() * 1e6).toString(36)}`;
}
export const emptyTree = (): FilterGroup => ({ id: newId('root'), type: 'group', combinator: 'and', rules: [] });

// ── Tree mutation helpers ───────────────────────────────────
// Recursive — walks plain groups AND the inner predicate of lambda nodes.
export function patchNode(
  root: FilterGroup,
  id: string,
  patch: Partial<Omit<FilterGroup, 'type'>> &
         Partial<Omit<FilterRule, 'type'>> &
         Partial<Omit<FilterFunctionNode, 'type'>> &
         Partial<Omit<FilterLambdaNode, 'type'>>,
): FilterGroup {
  const visit = (n: FilterNode): FilterNode => {
    if (n.id === id) {
      switch (n.type) {
        case 'group':    return { ...n, ...(patch as Partial<FilterGroup>) };
        case 'rule':     return { ...n, ...(patch as Partial<FilterRule>) };
        case 'function': return { ...n, ...(patch as Partial<FilterFunctionNode>) };
        case 'lambda':   return { ...n, ...(patch as Partial<FilterLambdaNode>) };
      }
    }
    if (n.type === 'group') return { ...n, rules: n.rules.map(visit) };
    if (n.type === 'lambda') return { ...n, inner: visit(n.inner) as FilterGroup };
    return n;
  };
  return visit(root) as FilterGroup;
}

export function addChild(root: FilterGroup, parentId: string, child: FilterNode): FilterGroup {
  const visit = (n: FilterNode): FilterNode => {
    if (n.type === 'group') {
      if (n.id === parentId) return { ...n, rules: [...n.rules, child] };
      return { ...n, rules: n.rules.map(visit) };
    }
    if (n.type === 'lambda') {
      return { ...n, inner: visit(n.inner) as FilterGroup };
    }
    return n;
  };
  return visit(root) as FilterGroup;
}

/**
 * Reorder the children of a specific group. Walks the tree (including
 * lambda inner-groups) to find the group with `parentId`, then moves its
 * `rules[fromIndex]` to `rules[toIndex]`. Returns a new tree; original
 * is not mutated. Used by the FilterEditor's DnD reorder UX.
 */
export function reorderInGroup(root: FilterGroup, parentId: string, fromIndex: number, toIndex: number): FilterGroup {
  const visit = (n: FilterNode): FilterNode => {
    if (n.type === 'group') {
      if (n.id === parentId) {
        if (fromIndex < 0 || fromIndex >= n.rules.length || toIndex < 0 || toIndex >= n.rules.length) return n;
        const next = [...n.rules];
        const [moved] = next.splice(fromIndex, 1);
        next.splice(toIndex, 0, moved);
        return { ...n, rules: next };
      }
      return { ...n, rules: n.rules.map(visit) };
    }
    if (n.type === 'lambda') return { ...n, inner: visit(n.inner) as FilterGroup };
    return n;
  };
  return visit(root) as FilterGroup;
}

export function removeNode(root: FilterGroup, id: string): FilterGroup {
  const visit = (n: FilterNode): FilterNode => {
    if (n.type === 'group') {
      return {
        ...n,
        rules: n.rules
          .filter(c => c.id !== id)
          .map(c => visit(c)),
      };
    }
    if (n.type === 'lambda') return { ...n, inner: visit(n.inner) as FilterGroup };
    return n;
  };
  return visit(root) as FilterGroup;
}

export function findNode(root: FilterGroup, id: string): FilterNode | null {
  const visit = (n: FilterNode): FilterNode | null => {
    if (n.id === id) return n;
    if (n.type === 'group') {
      for (const c of n.rules) {
        const m = visit(c);
        if (m) return m;
      }
    }
    if (n.type === 'lambda') return visit(n.inner);
    return null;
  };
  return visit(root);
}

export function countRules(g: FilterGroup): number {
  let n = 0;
  for (const c of g.rules) {
    switch (c.type) {
      case 'group':    n += countRules(c); break;
      case 'lambda':   n += countRules(c.inner) || 1; break; // empty lambda counts as one
      case 'rule':
      case 'function': n += 1; break;
    }
  }
  return n;
}

// ── OData encoding ──────────────────────────────────────────
function escapeStringLiteral(s: string): string {
  return s.replace(/'/g, "''");
}

/**
 * Strip leading wildcards from contains/startswith/endswith values.
 *
 * Per https://learn.microsoft.com/en-us/power-apps/developer/data-platform/wildcard-characters
 * Dataverse Web API recognizes these wildcards in string-function arguments:
 *   %      — matches zero+ characters (any position)
 *   _      — matches a single character (any position)
 *   [a-z]  — matches a character in the range
 *   [^abc] — matches a character NOT in the set
 *
 * Per https://learn.microsoft.com/en-us/power-apps/developer/data-platform/query-antipatterns
 * leading wildcards force a table scan and are *heavily throttled*. The doc
 * also calls out hyphen `-` and apostrophe `'` as collation-sensitive characters
 * that behave like leading wildcards (e.g. `-%`, `'%751`).
 *
 * This function returns the cleaned value AND a classification of what was
 * stripped so the caller can render a faithful advisory ("we removed the
 * leading % from your search — Dataverse doesn't support it").
 */
export type StrippedKind = '%' | '_' | '[' | "[^" | '-' | "'";

export function stripLeadingWildcards(raw: string): {
  value: string;
  stripped: boolean;
  /** Ordered list of leading-character kinds that were stripped (deduped). */
  kinds: StrippedKind[];
} {
  if (!raw) return { value: raw, stripped: false, kinds: [] };

  const kinds: StrippedKind[] = [];
  let i = 0;
  while (i < raw.length) {
    const ch = raw[i];
    // % and _ — classic SQL/OData wildcards
    if (ch === '%') { if (!kinds.includes('%')) kinds.push('%'); i++; continue; }
    if (ch === '_') { if (!kinds.includes('_')) kinds.push('_'); i++; continue; }
    // Hyphen / apostrophe — collation-driven leading-wildcard analogs
    if (ch === '-' && i === 0) { kinds.push('-'); i++; continue; }
    if (ch === "'" && i === 0) { kinds.push("'"); i++; continue; }
    // [ and [^ — bracket character class. Skip the whole [...] block.
    if (ch === '[') {
      const isNeg = raw[i + 1] === '^';
      kinds.push(isNeg ? '[^' : '[');
      const close = raw.indexOf(']', i);
      if (close === -1) {
        // Malformed bracket — strip the rest as a "[" leading wildcard.
        i = raw.length;
      } else {
        i = close + 1;
      }
      continue;
    }
    // First non-wildcard character — stop.
    break;
  }

  const value = raw.slice(i);
  return { value, stripped: i > 0, kinds };
}

/**
 * Walk a filter tree and return every contains/startswith/endswith rule whose
 * value begins with a leading wildcard (after the rule is encoded, that
 * wildcard is silently stripped — see ruleToOData). The rewrite needs to
 * be surfaced to the user; this scanner feeds the wildcard advisory.
 *
 * The op-name match is intentionally narrow — only the OData `odata-fn` ops
 * call stripLeadingWildcards in the encoder; comparison/null-check/lambda/
 * Dataverse-fn paths never see the strip.
 */
export interface StrippedWildcardEntry {
  col: string;
  raw: string;
  cleaned: string;
  kinds: StrippedKind[];
}

const STRING_FN_OPS = new Set(['contains', 'startswith', 'endswith']);

export function collectStrippedWildcards(group: FilterGroup, out: StrippedWildcardEntry[] = []): StrippedWildcardEntry[] {
  for (const node of group.rules) {
    if (node.type === 'rule' && STRING_FN_OPS.has(node.op) && node.val) {
      const { value, stripped, kinds } = stripLeadingWildcards(node.val);
      if (stripped) {
        out.push({ col: node.col, raw: node.val, cleaned: value, kinds });
      }
    } else if (node.type === 'group') {
      collectStrippedWildcards(node, out);
    } else if (node.type === 'lambda') {
      collectStrippedWildcards(node.inner, out);
    }
  }
  return out;
}

/**
 * Human-readable description of each leading-wildcard kind for advisory copy.
 */
export function strippedKindLabel(k: StrippedKind): string {
  switch (k) {
    case '%':  return 'percent (%) — matches any sequence';
    case '_':  return 'underscore (_) — matches a single character';
    case '[':  return 'bracket character class ([...])';
    case '[^': return 'negated bracket character class ([^...])';
    case '-':  return 'hyphen (-) — collation-driven leading wildcard';
    case "'":  return 'apostrophe (\') — collation-driven leading wildcard';
  }
}

function quoteLiteral(col: ColumnMeta | undefined, raw: string): string {
  if (raw === '' || raw == null) return "''";
  if (!col) return `'${escapeStringLiteral(raw)}'`;
  switch (col.attributeType) {
    case 'String':
    case 'Memo':
      return `'${escapeStringLiteral(raw)}'`;
    case 'DateTime':
      // ISO 8601 — no quotes for DateTimeOffset; DateOnly columns use a plain date literal.
      return raw.includes('T') || /^\d{4}-\d{2}-\d{2}/.test(raw) ? raw : `'${raw}'`;
    case 'Boolean':
      return raw === 'true' || raw === '1' ? 'true' : 'false';
    case 'Uniqueidentifier':
    case 'Lookup':
    case 'Customer':
    case 'Owner':
      return raw; // GUID literal — no quotes per OData v4
    case 'BigInt':
    case 'Integer':
    case 'Decimal':
    case 'Double':
    case 'Money':
    case 'Picklist':
    case 'MultiSelectPicklist':
    case 'State':
    case 'Status':
    case 'EntityName':
      return raw;
    case 'File':
    case 'Image':
      return raw; // shouldn't reach here — File/Image are hidden in the column picker
    default:
      return `'${escapeStringLiteral(raw)}'`;
  }
}
const arrayLiteral = (vals: string[]): string =>
  JSON.stringify(vals.map(v => String(v ?? '')));

function colLookup(col: string, prefix: string | undefined, table: TableMeta): ColumnMeta | undefined {
  // Resolve the leaf through the canonical metadata-driven resolver so a
  // nav-path leaf (`msdyn_opportunityid/abc_salesstage`) is looked up on the
  // RELATED entity, not the root. This is what makes the leaf's AttributeType
  // drive quoting in `quoteLiteral` — fixing the custom-OptionSet "quoted as
  // a string" bug (#33). `prefix` strips an enclosing lambda alias. For a bare
  // column the resolver behaves exactly like the old root-table lookup.
  //
  // Returns undefined while a hop's metadata is still loading; the warm layer
  // (`useWarmReferencedTables`) ensures referenced related entities are fetched
  // so this resolves correctly by encode time, then re-renders.
  return resolveNavPath(table, col, { alias: prefix }).leaf;
}

export interface EncodeCtx {
  table: TableMeta;
  /** Alias prefix used by enclosing lambda (e.g. `c`), if any */
  lambdaAlias?: string;
}

export function ruleToOData(rule: FilterRule, ctx: EncodeCtx): string {
  const op = findOperator(rule.op);
  if (!op) return '';
  const col = colLookup(rule.col, ctx.lambdaAlias, ctx.table);
  // Lookup columns address their VALUE at `_<logical>_value` in OData,
  // not at the bare logical name (which is a navigation property — only
  // valid in `$expand` and lambda paths). Preserve any alias / nav path
  // prefix and rewrite the trailing attribute segment via the shared
  // `attrRef` helper (driven by the column's pre-computed oDataName).
  const colExpr = (() => {
    const parts = rule.col.split('/');
    const last = parts.pop() ?? rule.col;
    const lastEnc = attrRef(col, last);
    return [...parts, lastEnc].join('/');
  })();
  // RHS — bare column name OR quoted literal per AttributeType
  const rhs = (raw: string): string => {
    if (rule.valKind === 'column' && raw) {
      // Bare property name on RHS — same row column reference.
      // If the LHS uses a lambda alias prefix (e.g. `c/jobtitle`) we
      // mirror that on the RHS so both sides resolve in the same scope.
      // Same lookup-value rewrite applies on the RHS.
      const rhsCol = ctx.table.columns.find(c => c.logicalName === raw);
      const rhsEnc = attrRef(rhsCol, raw);
      return ctx.lambdaAlias && !raw.includes('/') ? `${ctx.lambdaAlias}/${rhsEnc}` : rhsEnc;
    }
    return quoteLiteral(col, raw);
  };

  switch (op.kind) {
    case 'comparison':
      return `${colExpr} ${op.odata} ${rhs(rule.val ?? '')}`;
    case 'null-check':
      return `${colExpr} ${op.odata}`;
    case 'odata-fn': {
      // Leading wildcards are explicitly unsupported — strip them.
      const { value } = stripLeadingWildcards(rule.val ?? '');
      const expr = `${op.odata}(${colExpr},${quoteLiteral(col, value)})`;
      return rule.negated ? `not ${expr}` : expr;
    }
    default:
      return '';
  }
}

/**
 * Compute the `PropertyName=` string for a Dataverse query function. The rule
 * (verified empirically against the live Web API) is:
 *
 *   • For all DV functions, `PropertyName` is the BARE logical name. The
 *     `_<attr>_value` form returns 0x80041103 even on lookup columns (tests
 *     5.8 / 5.12 / 13.13). The `In`/`EqualUserId`/`EqualUserTeams`/… side
 *     match: bare name works (tests 5.9 / 13.12 / 13.14).
 *
 *   • Inside a lambda, the alias is NEVER prefixed in `PropertyName`. The
 *     function evaluates in the collection scope automatically (test 6.5b
 *     succeeds with `'createdon'`; test 6.5 fails with `'c/createdon'`).
 *
 *   • For hierarchy functions (`Above` / `Under` / …) the argument MUST be
 *     the entity's PRIMARY KEY column, not the parent lookup. Targeting
 *     `parentaccountid` returns 0x80040203 "is NOT a primary key" (test
 *     5.13). Targeting `accountid` succeeds (test 5.13b). Whether the entity
 *     actually has a hierarchical relationship configured is checked by
 *     Dataverse at runtime — `validateRequest` surfaces a hint advisory.
 */
function propertyNameForFunction(
  fn: FilterFunctionNode,
  op: ReturnType<typeof findOperator>,
  ctx: EncodeCtx,
): string {
  // 1. Strip any leading lambda-alias segment (`c/createdon` → `createdon`).
  let bare = fn.col;
  if (ctx.lambdaAlias && bare.startsWith(ctx.lambdaAlias + '/')) {
    bare = bare.slice(ctx.lambdaAlias.length + 1);
  }
  // 2. Strip an erroneous `_<x>_value` form if the rule was persisted with
  //    the wrong column shape (legacy state migration).
  if (bare.startsWith('_') && bare.endsWith('_value')) {
    bare = bare.slice(1, -'_value'.length);
  }
  // 3. Hierarchy functions: swap to the entity's primary key. The user's
  //    column pick (typically the self-referential lookup) is what defines
  //    the *category* of hierarchy, but Dataverse demands the PK for the
  //    `PropertyName` argument.
  if (op?.category === 'hierarchy') {
    return ctx.table.primaryKey || bare;
  }
  return bare;
}

export function functionToOData(fn: FilterFunctionNode, ctx: EncodeCtx): string {
  const op = findOperator(fn.op);
  if (!op) return '';
  const col = colLookup(fn.col, ctx.lambdaAlias, ctx.table);
  const propName = propertyNameForFunction(fn, op, ctx);

  let expr = '';
  switch (op.kind) {
    case 'dv-fn-0':
      expr = `${op.odata}(PropertyName='${propName}')`;
      break;
    case 'dv-fn-1': {
      const v = fn.val ?? '';
      const lit = col && (col.attributeType === 'DateTime' || col.attributeType === 'String' || col.attributeType === 'Memo')
        ? `'${escapeStringLiteral(v)}'` : v;
      expr = `${op.odata}(PropertyName='${propName}',PropertyValue=${lit})`;
      break;
    }
    case 'dv-fn-1-int':
      expr = `${op.odata}(PropertyName='${propName}',PropertyValue=${fn.val || '0'})`;
      break;
    case 'dv-fn-2': {
      const [a, b] = fn.vals ?? ['', ''];
      expr = `${op.odata}(PropertyName='${propName}',PropertyValue1=${a || '0'},PropertyValue2=${b || '0'})`;
      break;
    }
    case 'dv-fn-array': {
      // PropertyValues is a JSON array of STRINGS — even for GUIDs and ints.
      // Test 13.22 confirmed unquoted GUIDs are rejected ("Invalid JSON.
      // The value '98413e' is not a valid number."). Ints-as-strings are
      // accepted (test 5.3). `arrayLiteral` wraps every value with String()
      // and JSON.stringify, which gives us the right shape.
      const vals = fn.values ?? (fn.vals ? [...fn.vals] : []);
      expr = `${op.odata}(PropertyName='${propName}',PropertyValues=${arrayLiteral(vals)})`;
      break;
    }
    case 'dv-fn-guid':
      expr = `${op.odata}(PropertyName='${propName}',PropertyValue='${fn.val ?? ''}')`;
      break;
    default:
      return '';
  }
  // NOTE: we never emit `not` on a Dataverse query function. Dataverse
  // rejects `not Microsoft.Dynamics.CRM.<anything>` with 0x80060888 / 405
  // (test 5.7) — the negatable functions all expose explicit `Not*`
  // siblings (`NotBetween`, `NotIn`, `NotEqualUserId`, `DoesNotBeginWith`,
  // `DoesNotContainValues`, …) which the user picks via the operator
  // combobox instead. The `fn.negated` field is kept on the type so
  // legacy persisted requests round-trip without data loss, but the
  // encoder ignores it here.
  return expr;
}

export function lambdaToOData(
  node: FilterLambdaNode,
  parentTable: TableMeta,
  /**
   * Alias of the ENCLOSING lambda scope, if this lambda is nested.
   * Per OData, every column / nav reference inside a lambda predicate
   * must address the iteration variable — including a NESTED lambda's
   * own `<nav>/`. Example:
   *   contact_customer_accounts/any(c: c/Contact_Annotation/any(n: n/notetext ne null))
   *                                    ^^                    ^^^^^^^^^^^^^^^^^^^^
   *                            outer alias prefix      inner lambda body (already prefixed by ruleToOData)
   * Without this prefix, Dataverse rejects with: "Conditions on property
   * other than current navigation property in any/all is not supported."
   */
  outerAlias?: string,
): string {
  // Resolve the related table for inner column lookups
  const nav = parentTable.navigationProperties.find(n => n.name === node.nav);
  const innerTbl = nav ? findTableByLogical(nav.targetEntity) : undefined;
  const inner = innerTbl
    ? groupToOData(node.inner, innerTbl, node.alias)
    : groupToOData(node.inner, parentTable, node.alias);

  const navPath = outerAlias ? `${outerAlias}/${node.nav}` : node.nav;
  const body = !inner
    ? `${navPath}/${node.lambda}()`
    : `${navPath}/${node.lambda}(${node.alias}:${inner})`;
  // `not <nav>/any(...)` and `not <nav>/all(...)` are both accepted by
  // Dataverse (test G.8). No parens required around the lambda body —
  // `not` binds to the next boolean-producing primary expression.
  return node.negated ? `not ${body}` : body;
}

export function groupToOData(g: FilterGroup, table: TableMeta, lambdaAlias?: string): string {
  const ctx: EncodeCtx = { table, lambdaAlias };
  const parts = g.rules
    .map(c => {
      switch (c.type) {
        case 'group':    return groupToOData(c, table, lambdaAlias);
        case 'rule':     return ruleToOData(c, ctx);
        case 'function': return functionToOData(c, ctx);
        // Propagate the current scope's alias so a nested lambda's nav
        // gets prefixed (e.g. `c/Contact_Annotation/any(n: …)`).
        case 'lambda':   return lambdaToOData(c, table, lambdaAlias);
      }
    })
    .filter(Boolean);
  if (parts.length === 0) return '';
  let body: string;
  if (parts.length === 1) {
    body = parts[0];
  } else {
    const joiner = ` ${g.combinator} `;
    body = parts
      .map(p => (p.includes(' and ') || p.includes(' or ')) ? `(${p})` : p)
      .join(joiner);
  }
  // Group-level NOT — wrap the body in `not (…)`. Empirically supported
  // (tests G.1–G.7). Parens are required even for a single-child group
  // because `not` has higher precedence than the comparison operators in
  // OData v4 — `not name eq 'x'` would parse as `(not name) eq 'x'`.
  // Encoder is faithful regardless of body content; validateRequest
  // surfaces a warning when the negated group contains a DV function
  // (test G.9 — Dataverse rejects `not (<dv-fn>)` with 405).
  if (g.negated) {
    return `not (${body})`;
  }
  return body;
}

// Pulled in lazily to avoid circular deps at module-init time.
let _tableLookup: ((id: string) => TableMeta | undefined) | null = null;
export function _setTableLookup(fn: (id: string) => TableMeta | undefined): void {
  _tableLookup = fn;
}
function findTableByLogical(id: string): TableMeta | undefined {
  return _tableLookup ? _tableLookup(id) : undefined;
}

// ── Validators ──────────────────────────────────────────────
export function validateTree(g: FilterGroup, urlBytes: number): {
  ruleCount: number;
  ruleWarn: boolean;
  ruleError: boolean;
  urlWarn: boolean;
  urlError: boolean;
} {
  const ruleCount = countRules(g);
  return {
    ruleCount,
    ruleWarn:  ruleCount >= 400,
    ruleError: ruleCount  > 500,
    urlWarn:   urlBytes  >= 30_000,
    urlError:  urlBytes  >= 32_000,
  };
}

// Factory helpers for the "Add …" menu in the FilterEditor footer.
export function makeRuleDefaults(op: OperatorDef): Partial<FilterRule> {
  switch (op.kind) {
    case 'null-check':       return {};
    default:                 return { val: '' };
  }
}

export function defaultRule(firstColumn: string): FilterRule {
  return { id: newId('r'), type: 'rule', col: firstColumn, op: 'eq', val: '' };
}

export function defaultGroup(firstColumn: string): FilterGroup {
  return { id: newId('g'), type: 'group', combinator: 'or', rules: [defaultRule(firstColumn)] };
}

/**
 * Build a default function node. Picks a column on `table` whose AttributeType
 * is in the operator's allowedTypes — e.g. `LastXDays` (date-only) gets
 * defaulted to the first DateTime column rather than a wrong-type column.
 * Falls back to the first non-hidden column if none match.
 */
export function defaultFunction(table: TableMeta, fnId = 'LastXDays'): FilterFunctionNode {
  const op = findOperator(fnId);
  const allowed = op?.allowedTypes;
  const visibleCols = table.columns.filter(c => c.attributeType !== 'File' && c.attributeType !== 'Image');
  const compatible = allowed
    ? visibleCols.find(c => allowed.includes(c.attributeType))
    : visibleCols[0];
  const col = compatible?.logicalName ?? visibleCols[0]?.logicalName ?? table.columns[0].logicalName;
  // Sensible default value for int-valued rolling functions
  const val = op?.intValue ? '30' : '';
  return { id: newId('f'), type: 'function', op: fnId, col, val };
}

export function defaultLambda(navName: string, firstColumn: string): FilterLambdaNode {
  // Use the first letter of the nav as the alias (e.g. contact_customer_accounts → c)
  const alias = navName[0] ?? 'x';
  return {
    id: newId('l'),
    type: 'lambda',
    nav: navName,
    lambda: 'any',
    alias,
    inner: { id: newId('g'), type: 'group', combinator: 'and', rules: [
      { id: newId('r'), type: 'rule', col: `${alias}/${firstColumn}`, op: 'eq', val: '' },
    ] },
  };
}
