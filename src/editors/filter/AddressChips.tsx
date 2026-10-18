// Address shortcuts for the $filter builder (#7).
//
// Tables with address composite columns (account, contact, lead, …) carry
// `address1_city`, `address1_country`, `address2_postalcode`, … Filtering by
// them means finding the right one in a long column list, so the builder
// offers one-click chips that add the rule with the column and operator set:
// "City = …", "State / Province = …", "Country = …", "Postal code starts
// with …". The user then types the value in the new rule.

import { Button, Caption1, Tooltip, tokens } from '@fluentui/react-components';
import { Add16Regular, Location16Regular } from '@fluentui/react-icons';
import type { ColumnMeta, TableMeta } from '../../mock/metadata';
import { defaultRule, type FilterRule } from './filterTree';

const ADDRESS_PREFIXES = ['address1', 'address2', 'address3'] as const;

const ADDRESS_FIELDS: { suffix: string; label: string; op: 'eq' | 'startswith' }[] = [
  { suffix: 'city', label: 'City =', op: 'eq' },
  { suffix: 'stateorprovince', label: 'State / Province =', op: 'eq' },
  { suffix: 'country', label: 'Country =', op: 'eq' },
  { suffix: 'postalcode', label: 'Postal code starts with', op: 'startswith' },
];

export interface AddressShortcut {
  /** Column logical name, e.g. `address1_city`. */
  col: string;
  op: 'eq' | 'startswith';
  label: string;
  /** Column display name, for the tooltip. */
  displayName: string;
}

export interface AddressShortcutGroup {
  prefix: string;
  /** `Address 1`, `Address 2`, … */
  label: string;
  shortcuts: AddressShortcut[];
}

function filterableText(c: ColumnMeta | undefined): c is ColumnMeta {
  return (
    !!c &&
    (c.attributeType === 'String' || c.attributeType === 'Memo') &&
    c.isValidForRead !== false &&
    c.isValidOData !== false
  );
}

/** The address shortcuts a table supports, one group per address (empty when none). */
export function addressShortcuts(table: TableMeta): AddressShortcutGroup[] {
  const byName = new Map(table.columns.map((c) => [c.logicalName, c]));
  const groups: AddressShortcutGroup[] = [];
  for (const prefix of ADDRESS_PREFIXES) {
    const shortcuts: AddressShortcut[] = [];
    for (const f of ADDRESS_FIELDS) {
      const col = byName.get(`${prefix}_${f.suffix}`);
      if (!filterableText(col)) continue;
      shortcuts.push({
        col: col.logicalName,
        op: f.op,
        label: f.label,
        displayName: col.displayName,
      });
    }
    if (shortcuts.length) {
      groups.push({ prefix, label: `Address ${prefix.slice(-1)}`, shortcuts });
    }
  }
  return groups;
}

/** A new rule for a shortcut: column + operator set, value left for the user. */
export function addressRule(s: AddressShortcut): FilterRule {
  return { ...defaultRule(s.col), op: s.op, val: '' };
}

export function AddressChips({
  table,
  onAdd,
}: {
  table: TableMeta;
  onAdd: (rule: FilterRule) => void;
}) {
  const groups = addressShortcuts(table);
  if (!groups.length) return null;
  return (
    <div
      role="group"
      aria-label="Address filter shortcuts"
      style={{ display: 'flex', flexDirection: 'column', gap: 6, marginBottom: 12 }}
    >
      {groups.map((g) => (
        <div
          key={g.prefix}
          style={{ display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' }}
        >
          <Caption1
            style={{
              display: 'inline-flex',
              alignItems: 'center',
              gap: 4,
              minWidth: 84,
              color: tokens.colorNeutralForeground3,
            }}
          >
            <Location16Regular />
            {g.label}
          </Caption1>
          {g.shortcuts.map((s) => (
            <Tooltip
              key={s.col}
              content={`Add a rule: ${s.displayName} (${s.col}) ${s.op === 'eq' ? 'equals' : 'starts with'} …`}
              relationship="description"
            >
              <Button
                size="small"
                shape="circular"
                appearance="outline"
                icon={<Add16Regular />}
                onClick={() => onAdd(addressRule(s))}
              >
                {s.label} …
              </Button>
            </Tooltip>
          ))}
        </div>
      ))}
    </div>
  );
}
