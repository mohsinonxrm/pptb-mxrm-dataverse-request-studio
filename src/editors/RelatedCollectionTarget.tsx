// Predefined Query → "Related collection" target (#5).
//
// A saved / user query can run against any collection of its table's type,
// including a collection-valued nav on one row:
//   GET /accounts(<id>)/opportunity_parent_account?savedQuery=<id>
// https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/retrieve-and-execute-predefined-queries#apply-a-query-to-any-collection-of-the-appropriate-type
//
// The relationship list comes from the QUERY table's own navs: each lookup
// (many-to-one) or many-to-many nav on it is one side of a relationship
// whose other side is a collection-valued nav on the source table. Picking
// a relationship loads the source table and resolves that nav by
// relationship name; the user then picks the source row.

import { useEffect, useMemo } from 'react';
import { Caption1, Dropdown, Field, Option, Spinner, tokens } from '@fluentui/react-components';
import { findTable, type TableMeta } from '../mock/metadata';
import { useLiveTable } from '../host/useLiveMetadata';
import { RecordPicker } from '../primitives/RecordPicker';
import type { PredefinedRelatedTarget } from '../state/readState';

interface RelationshipOption {
  relationshipName: string;
  sourceTable: string;
  /** How the query table points back — the lookup column or M:N nav. */
  via: string;
  manyToMany: boolean;
}

/** Relationships through which rows of `table` form a collection on another row. */
export function relatedCollectionOptions(table: TableMeta): RelationshipOption[] {
  const seen = new Set<string>();
  const out: RelationshipOption[] = [];
  for (const n of table.navigationProperties) {
    if (n.cardinality === 'OneToMany') continue;
    const key = `${n.relationshipName}|${n.targetEntity}`;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push({
      relationshipName: n.relationshipName,
      sourceTable: n.targetEntity,
      via: n.referencingAttribute ?? n.name,
      manyToMany: n.cardinality === 'ManyToMany',
    });
  }
  return out.sort(
    (a, b) =>
      a.sourceTable.localeCompare(b.sourceTable) ||
      a.relationshipName.localeCompare(b.relationshipName),
  );
}

/** The source table's collection-valued nav for a relationship, targeting `table`. */
export function sourceNavFor(
  source: TableMeta,
  relationshipName: string,
  table: string,
): string | null {
  const nav = source.navigationProperties.find(
    (n) =>
      n.relationshipName === relationshipName &&
      n.cardinality !== 'ManyToOne' &&
      n.targetEntity === table,
  );
  return nav?.name ?? null;
}

export function RelatedCollectionTarget({
  table,
  value,
  onChange,
}: {
  /** Query table (logical name) — its `returnedtypecode`. */
  table: TableMeta;
  value: PredefinedRelatedTarget;
  onChange: (next: PredefinedRelatedTarget) => void;
}) {
  const options = useMemo(() => relatedCollectionOptions(table), [table]);
  const { table: source, loading } = useLiveTable(value.sourceTable || null);

  // Resolve the source-side nav once the source table's metadata lands.
  useEffect(() => {
    if (!source || !value.relationshipName) return;
    const nav = sourceNavFor(source, value.relationshipName, table.logicalName) ?? '';
    if (nav !== value.navProperty) onChange({ ...value, navProperty: nav });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [source, value.relationshipName, table.logicalName]);

  const key = (o: RelationshipOption) => `${o.relationshipName}|${o.sourceTable}`;
  const selected = options.find(
    (o) => o.relationshipName === value.relationshipName && o.sourceTable === value.sourceTable,
  );
  const label = (o: RelationshipOption) =>
    `${findTable(o.sourceTable)?.displayName ?? o.sourceTable} · ${o.relationshipName}`;

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
      <Field
        label="Relationship"
        hint={
          options.length
            ? `Rows of ${table.displayName} related to one row of another table.`
            : undefined
        }
        validationState={options.length ? 'none' : 'warning'}
        validationMessage={
          options.length
            ? undefined
            : `${table.displayName} has no lookups or many-to-many relationships.`
        }
      >
        <Dropdown
          placeholder="Pick a relationship…"
          value={selected ? label(selected) : ''}
          selectedOptions={selected ? [key(selected)] : []}
          onOptionSelect={(_, d) => {
            const o = options.find((x) => key(x) === d.optionValue);
            if (!o) return;
            onChange({
              sourceTable: o.sourceTable,
              relationshipName: o.relationshipName,
              navProperty: '',
              recordId: o.sourceTable === value.sourceTable ? value.recordId : null,
              recordName: o.sourceTable === value.sourceTable ? value.recordName : undefined,
            });
          }}
          listbox={{ style: { maxHeight: 360 } }}
        >
          {options.map((o) => (
            <Option key={key(o)} value={key(o)} text={label(o)}>
              <div style={{ display: 'flex', flexDirection: 'column', minWidth: 0 }}>
                <span>{label(o)}</span>
                <Caption1
                  style={{
                    color: tokens.colorNeutralForeground3,
                    fontFamily: tokens.fontFamilyMonospace,
                  }}
                >
                  {o.manyToMany ? `many-to-many · ${o.via}` : `via ${o.via}`}
                </Caption1>
              </div>
            </Option>
          ))}
        </Dropdown>
      </Field>

      {value.sourceTable && loading && !source && (
        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <Spinner size="tiny" />
          <Caption1>Loading {value.sourceTable} metadata…</Caption1>
        </div>
      )}

      {source && value.relationshipName && !value.navProperty && (
        <Caption1 style={{ color: tokens.colorPaletteRedForeground1 }}>
          {source.displayName} has no collection navigation property for {value.relationshipName}.
        </Caption1>
      )}

      {source && value.navProperty && (
        <Field
          label={`${source.displayName} row`}
          hint={
            <span>
              Collection: <code>{`${source.entitySetName}(…)/${value.navProperty}`}</code>
            </span>
          }
        >
          <RecordPicker
            table={source.logicalName}
            selectedId={value.recordId}
            onPick={(rec) =>
              onChange({ ...value, recordId: rec?.id ?? null, recordName: rec?.primary })
            }
          />
        </Field>
      )}
    </div>
  );
}
