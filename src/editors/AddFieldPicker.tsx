import { useState } from 'react';
import { Caption1, Combobox, Option, tokens } from '@fluentui/react-components';
import type { ColumnMeta } from '../metadata/metadata';
import { MoreOptionsHint, useOptionSearch } from '../primitives/optionSearch';

/**
 * FieldSetEditor's "Add field" picker. Type-to-filter over the addable
 * columns with a capped option list: a wide table can offer 500+ columns,
 * and Fluent renders every Option as soon as the input gains focus (#21).
 */
export function AddFieldPicker({
  addable,
  onPick,
}: {
  addable: ColumnMeta[];
  onPick: (logicalName: string) => void;
}) {
  const [query, setQuery] = useState('');
  const { visible, matchCount, hiddenCount } = useOptionSearch(addable, query, (c) => [
    c.displayName,
    c.logicalName,
  ]);
  return (
    <Combobox
      freeform
      size="small"
      value={query}
      placeholder={`Add one of ${addable.length} fields…`}
      style={{ minWidth: 320 }}
      onChange={(e) => setQuery((e.target as HTMLInputElement).value)}
      onOptionSelect={(_, d) => {
        if (!d.optionValue) return;
        onPick(d.optionValue);
      }}
    >
      {matchCount === 0 && (
        <Option value="__none" text="" disabled>
          <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
            No fields match "{query}"
          </Caption1>
        </Option>
      )}
      {visible.map((c) => (
        <Option key={c.logicalName} value={c.logicalName} text={c.displayName}>
          <div style={{ display: 'flex', flexDirection: 'column' }}>
            <span>{c.displayName}</span>
            <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
              {c.logicalName} · <code>{c.attributeType}</code>
              {c.required && (
                <span style={{ marginLeft: 6, color: tokens.colorPaletteRedForeground1 }}>
                  required
                </span>
              )}
            </Caption1>
          </div>
        </Option>
      ))}
      <MoreOptionsHint hiddenCount={hiddenCount} />
    </Combobox>
  );
}
