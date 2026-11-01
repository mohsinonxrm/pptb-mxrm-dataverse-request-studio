// JSON Schema for the field set's JSON view (#268). Monaco's JSON language
// service reads it to complete column names, `<nav>@odata.bind` keys and
// option values, to show each key's column on hover, and to underline a
// value of the wrong JSON type.
//
// It describes what `decodeFieldBody` maps, from the same inputs: the
// columns the form can set (`writable`) and the same JSON type per attribute
// type (NUMBER_TYPES, …), so a choice value outside its options isn't
// flagged either. One exception: a multi-select given as an array maps, but
// is underlined, because a verbatim body sends it as typed and Dataverse
// wants the comma-separated string. Other keys (annotations, unknown
// columns) aren't flagged here: the pane already lists them, and a verbatim
// body may need them. Kept free of Monaco imports, so it stays out of the
// Monaco chunk.

import {
  columnOptions,
  isLookupLike,
  type ColumnMeta,
  type NavProperty,
  type TableMeta,
} from '../metadata/metadata';
import { INT_CHOICE_TYPES, NUMBER_TYPES, STRING_TYPES } from './fieldBody';

/** The subset of JSON Schema (draft 7 + Monaco's extensions) the body uses. */
export interface BodyJsonSchema {
  type?: string | string[];
  description?: string;
  /** Monaco: value completions, with snippet placeholders. */
  defaultSnippets?: { label: string; description?: string; body: unknown }[];
  pattern?: string;
  /** Monaco: the warning shown when `pattern` doesn't match. */
  patternErrorMessage?: string;
  /** Monaco: validated and shown on hover, but not offered in completions. */
  doNotSuggest?: boolean;
  properties?: Record<string, BodyJsonSchema>;
}

export interface BodySchemaOptions {
  /** False when the field set has no explicit-null support (nulls are reported). */
  nullSupported: boolean;
  /** Offer the column in completions; the pickers' scope (#44). Default: all. */
  suggest?: (c: ColumnMeta) => boolean;
  /** A table's entity set, when known — for `@odata.bind` value snippets. */
  entitySetOf?: (logicalName: string) => string | undefined;
}

const EMPTY_GUID = '00000000-0000-0000-0000-000000000000';

/** The JSON types `decodeFieldBody` maps for a non-lookup column. */
function jsonTypesOf(c: ColumnMeta): string[] | undefined {
  const t = c.attributeType;
  if (INT_CHOICE_TYPES.has(t)) return ['integer'];
  if (NUMBER_TYPES.has(t)) return ['number'];
  if (t === 'Boolean') return ['boolean'];
  if (t === 'MultiSelectPicklist') return ['string'];
  if (STRING_TYPES.has(t)) return ['string'];
  if (t === 'EntityName') return ['string', 'number'];
  return undefined;
}

function requiredNote(c: ColumnMeta): string {
  if (c.requiredLevel === 'ApplicationRequired') return ' Business required.';
  if (c.requiredLevel === 'SystemRequired') return ' System required.';
  return '';
}

const OPTIONS_LISTED = 25;

/** "1 = Accounting, 2 = Agriculture", capped for long option sets. */
function optionList(options: { value: unknown; label: string }[]): string {
  const shown = options.slice(0, OPTIONS_LISTED).map((o) => `${String(o.value)} = ${o.label}`);
  const more = options.length - shown.length;
  return shown.join(', ') + (more > 0 ? `, and ${more} more` : '');
}

function columnSchema(c: ColumnMeta, nullable: boolean): BodyJsonSchema {
  const types = jsonTypesOf(c);
  const options = columnOptions(c);
  let description = `${c.displayName} (${c.attributeType}).${requiredNote(c)}`;
  if (c.attributeType === 'DateTime') {
    // Date Only behavior is Edm.Date in the Web API: no time part
    description +=
      c.dateTimeBehavior === 'DateOnly'
        ? ' A date, e.g. "2026-11-01".'
        : ' ISO 8601, e.g. "2026-11-01T09:00:00Z".';
  }
  if (c.attributeType === 'Boolean') {
    description += ` true = ${c.trueOption.label}, false = ${c.falseOption.label}.`;
  } else if (c.attributeType === 'MultiSelectPicklist' && options?.length) {
    description += ` Comma-separated values: ${optionList(options)}.`;
  } else if (INT_CHOICE_TYPES.has(c.attributeType) && options?.length) {
    description += ` Values: ${optionList(options)}.`;
  }
  if (c.description) description += `\n\n${c.description}`;

  const schema: BodyJsonSchema = { description };
  if (types) schema.type = nullable ? [...types, 'null'] : types.length === 1 ? types[0] : types;

  // Choices complete to their values, labelled in the list. Not an `enum`:
  // Monaco would list bare numbers (the labels only in the details pane),
  // and the decoder maps any integer, as Dataverse reports a wrong one.
  // Booleans get true / false from their type.
  if (INT_CHOICE_TYPES.has(c.attributeType) && options?.length) {
    schema.defaultSnippets = options.map((o) => ({
      label: `${o.value} · ${o.label}`,
      description: o.label,
      body: o.value,
    }));
  }
  if (c.attributeType === 'MultiSelectPicklist') {
    // The common shapes the decoder maps: integers between commas (it skips
    // empty pieces)
    schema.pattern = '^\\s*(-?\\d+\\s*)?(,\\s*(-?\\d+\\s*)?)*$';
    schema.patternErrorMessage = 'Expected comma-separated option values, e.g. "1,2".';
    if (options?.length) {
      schema.defaultSnippets = options.map((o) => ({
        label: `"${o.value}" · ${o.label}`,
        description: `${o.label}. Add more values after a comma.`,
        body: String(o.value),
      }));
    }
  }
  return schema;
}

/**
 * The `@odata.bind` keys for a lookup column, as `bindPropertyFor` names
 * them: one per ManyToOne navigation property that references the column
 * (polymorphic lookups have one per target), and `<column>@odata.bind` for
 * targets without one (all of them when the metadata has no navs).
 */
function bindKeysOf(
  tbl: TableMeta,
  c: ColumnMeta & { targets: string[] },
): { key: string; targets: string[] }[] {
  const navs = tbl.navigationProperties.filter(
    (n: NavProperty) => n.cardinality === 'ManyToOne' && n.referencingAttribute === c.logicalName,
  );
  const out = new Map<string, string[]>();
  const add = (key: string, targets: string[]) =>
    out.set(key, [...new Set([...(out.get(key) ?? []), ...targets])]);
  for (const n of navs) {
    // An Owner column's single nav targets `owner`; its rows are users or teams
    add(`${n.name}@odata.bind`, c.targets.includes(n.targetEntity) ? [n.targetEntity] : c.targets);
  }
  const covered = new Set([...out.values()].flat());
  const rest = c.targets.filter((t) => !covered.has(t));
  if (rest.length > 0 || out.size === 0) {
    add(`${c.logicalName}@odata.bind`, rest.length > 0 ? rest : c.targets);
  }
  return [...out].map(([key, targets]) => ({ key, targets }));
}

function bindSchema(
  c: ColumnMeta,
  targets: string[],
  nullable: boolean,
  entitySetOf: (logicalName: string) => string | undefined,
): BodyJsonSchema {
  const sets = targets.map((t) => ({ target: t, set: entitySetOf(t) }));
  const shown = sets.map((s) => `/${s.set ?? `<${s.target} entity set>`}(<guid>)`).join(' or ');
  return {
    type: nullable ? ['string', 'null'] : 'string',
    description:
      `${c.displayName} (${c.attributeType} → ${targets.join(', ')}).${requiredNote(c)} ` +
      `Value: "${shown}"${nullable ? ', or null to clear it' : ''}.`,
    defaultSnippets: sets.map((s) => ({
      label: `/${s.set ?? s.target}(<guid>)`,
      description: `A ${s.target} row`,
      body: s.set ? `/${s.set}(\${1:${EMPTY_GUID}})` : `/\${1:${s.target}s}(\${2:${EMPTY_GUID}})`,
    })),
  };
}

/** The JSON Schema for a body of `tbl` that sets `writable` columns. */
export function buildBodySchema(
  tbl: TableMeta,
  writable: readonly ColumnMeta[],
  { nullSupported, suggest = () => true, entitySetOf = () => undefined }: BodySchemaOptions,
): BodyJsonSchema {
  const properties: Record<string, BodyJsonSchema> = {};
  for (const c of writable) {
    const hidden = !suggest(c);
    if (isLookupLike(c)) {
      for (const { key, targets } of bindKeysOf(tbl, c)) {
        properties[key] = bindSchema(c, targets, nullSupported, entitySetOf);
        if (hidden) properties[key].doNotSuggest = true;
      }
      continue;
    }
    properties[c.logicalName] = columnSchema(c, nullSupported);
    if (hidden) properties[c.logicalName].doNotSuggest = true;
  }
  return { type: 'object', properties };
}
