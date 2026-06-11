// AlternateKeyEditor — picks the addressing strategy for an Upsert request.
//
// Modes:
//   • GUID            — same record-picker UX as Update (primary key)
//   • Alternate key   — pick a defined AlternateKeyDef and fill in the key
//                       columns. The URL builder slots the values into
//                       /<entitySet>(col1=val1,col2=val2).
//
// Per docs (use-upsert-insert-update-record):
//   "When using alternate keys, don't include the alternate key values in
//    the body of the request" — this editor only controls the URL key
//    segment; field values are still in the Field set pane.

import {
  Field,
  RadioGroup,
  Radio,
  Combobox,
  Option,
  Input,
  Caption1,
  tokens,
  MessageBar,
  MessageBarBody,
} from '@fluentui/react-components';
import { Key20Filled } from '@fluentui/react-icons';
import { PaneHead } from './PaneHead';
import { RecordPicker } from '../primitives/RecordPicker';
import { findTable, findColumn } from '../mock/metadata';
import type { UpsertKeyMode } from '../state/writeState';
import type { RequestGroup } from '../registry/requestTypes';

export interface AlternateKeyEditorProps {
  table: string;
  keyMode: UpsertKeyMode;
  setKeyMode: (k: UpsertKeyMode) => void;
  group?: RequestGroup;
}

export function AlternateKeyEditor({
  table,
  keyMode,
  setKeyMode,
  group = 'write',
}: AlternateKeyEditorProps) {
  const tbl = findTable(table);
  const altKeys = tbl?.alternateKeys ?? [];

  if (!tbl) {
    return (
      <MessageBar layout="multiline" intent="error">
        <MessageBarBody>
          Unknown table <code>{table}</code>.
        </MessageBarBody>
      </MessageBar>
    );
  }

  return (
    <div>
      <PaneHead
        icon={Key20Filled}
        title="Upsert key"
        sub="Identify the row to upsert — by GUID or by an alternate (business) key configured on the table."
        group={group}
      />

      <div style={{ maxWidth: 760 }}>
        <Field label="Addressing">
          <RadioGroup
            value={keyMode.kind}
            onChange={(_, d) => {
              if (d.value === 'guid') {
                setKeyMode({
                  kind: 'guid',
                  recordId: keyMode.kind === 'guid' ? keyMode.recordId : null,
                });
              } else {
                const first = altKeys[0];
                if (!first) return;
                const cur = keyMode.kind === 'alternate' ? keyMode.keyValues : {};
                setKeyMode({ kind: 'alternate', keyName: first.name, keyValues: cur });
              }
            }}
          >
            <Radio
              value="guid"
              label={
                <span>
                  <strong>By GUID</strong>
                  <Caption1
                    style={{
                      display: 'block',
                      color: tokens.colorNeutralForeground3,
                      marginTop: 2,
                    }}
                  >
                    Address the record by its primary-key (system) identifier — same as Update.
                  </Caption1>
                </span>
              }
            />
            <Radio
              value="alternate"
              disabled={altKeys.length === 0}
              label={
                <span>
                  <strong>By alternate key</strong>
                  {altKeys.length === 0 && (
                    <span style={{ color: tokens.colorNeutralForeground3, fontSize: 11 }}>
                      {' '}
                      · none defined on this table
                    </span>
                  )}
                  <Caption1
                    style={{
                      display: 'block',
                      color: tokens.colorNeutralForeground3,
                      marginTop: 2,
                    }}
                  >
                    Use a business identifier (e.g. account number, email) instead of the GUID. The
                    URL becomes{' '}
                    <code style={{ fontFamily: tokens.fontFamilyMonospace }}>
                      /{tbl.entitySetName}(col=value)
                    </code>
                    .
                  </Caption1>
                </span>
              }
            />
          </RadioGroup>
        </Field>

        {keyMode.kind === 'guid' && (
          <div style={{ marginTop: 14, maxWidth: 480 }}>
            <Field label="Record">
              <RecordPicker
                table={table}
                selectedId={keyMode.recordId}
                onPick={(r) => setKeyMode({ kind: 'guid', recordId: r?.id ?? null })}
                placeholder={`Search ${tbl.displayName} records…`}
              />
            </Field>
            <Caption1
              style={{ marginTop: 6, color: tokens.colorNeutralForeground3, display: 'block' }}
            >
              Leave empty to address by a future GUID — the URL will read <code>(&lt;id&gt;)</code>{' '}
              as a placeholder until you pick a row.
            </Caption1>
          </div>
        )}

        {keyMode.kind === 'alternate' && (
          <div
            style={{
              marginTop: 14,
              display: 'flex',
              flexDirection: 'column',
              gap: 14,
              maxWidth: 560,
            }}
          >
            <Field label="Key definition">
              <Combobox
                value={
                  altKeys.find((k) => k.name === keyMode.keyName)?.displayName ?? keyMode.keyName
                }
                selectedOptions={[keyMode.keyName]}
                onOptionSelect={(_, d) => {
                  if (!d.optionValue) return;
                  setKeyMode({ kind: 'alternate', keyName: d.optionValue, keyValues: {} });
                }}
              >
                {altKeys.map((k) => (
                  <Option key={k.name} value={k.name} text={k.displayName}>
                    <div style={{ display: 'flex', flexDirection: 'column' }}>
                      <span>{k.displayName}</span>
                      <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                        cols: {k.columns.join(', ')}
                      </Caption1>
                    </div>
                  </Option>
                ))}
              </Combobox>
            </Field>

            {(() => {
              const def = altKeys.find((k) => k.name === keyMode.keyName);
              if (!def) return null;
              return (
                <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                  {def.columns.map((c) => {
                    const col = findColumn(tbl, c);
                    return (
                      <Field
                        key={c}
                        label={col?.displayName ?? c}
                        hint={
                          <span>
                            <code style={{ fontFamily: tokens.fontFamilyMonospace }}>{c}</code>
                            {col && (
                              <>
                                {' '}
                                ·{' '}
                                <span style={{ color: tokens.colorBrandForeground2 }}>
                                  {col.attributeType}
                                </span>
                              </>
                            )}
                          </span>
                        }
                      >
                        <Input
                          size="small"
                          value={keyMode.keyValues[c] ?? ''}
                          onChange={(_, d) =>
                            setKeyMode({
                              kind: 'alternate',
                              keyName: keyMode.keyName,
                              keyValues: { ...keyMode.keyValues, [c]: d.value },
                            })
                          }
                          placeholder={col?.attributeType === 'String' ? 'e.g. ACC-0001' : 'value'}
                          style={{ fontFamily: tokens.fontFamilyMonospace }}
                        />
                      </Field>
                    );
                  })}
                </div>
              );
            })()}

            <MessageBar layout="multiline" intent="info">
              <MessageBarBody>
                When the row is <strong>created</strong> via upsert, Dataverse auto-sets these
                alternate-key columns from the URL —{' '}
                <strong>don't repeat them in the Field set</strong>. When the row is{' '}
                <strong>updated</strong>, alternate-key values in the URL identify the row and are
                ignored in the body.
              </MessageBarBody>
            </MessageBar>
          </div>
        )}
      </div>
    </div>
  );
}
