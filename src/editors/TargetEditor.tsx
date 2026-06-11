import { useEffect, useMemo, useState } from 'react';
import {
  Field,
  Combobox,
  Option,
  Caption1,
  tokens,
  Body1,
  Spinner,
  Button,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  MessageBarActions,
  type ComboboxProps,
} from '@fluentui/react-components';
import {
  Table20Filled,
  ArrowReset20Regular,
  ArrowSync20Regular,
  Settings20Regular,
} from '@fluentui/react-icons';
import { PaneHead } from './PaneHead';
import { findTable } from '../mock/metadata';
import { useLiveTable } from '../host/useLiveMetadata';
import { useScopedEntities } from '../host/useScopedEntities';
import { useOpenSettings } from '../host/useOpenSettings';
import { RecordPicker } from '../primitives/RecordPicker';
import { PasteODataDialog } from './PasteODataDialog';
import { metadata } from '../host/metadataProvider';
import type { RequestGroup } from '../registry/requestTypes';
import type { ParsedRequest } from '../engine/odataParser';

export interface TargetEditorProps {
  table: string;
  onTableChange: (logicalName: string) => void;
  /** If set, shows a record picker too (Retrieve Single) */
  recordId?: string | null;
  /**
   * Called when the user picks (or clears) a record.
   *
   * `primary` is the resolved primary-name value of the picked record
   * when known. For Search-mode picks it's always populated from the
   * typeahead result. For GUID-paste picks (where no name has been
   * resolved yet) it's an empty string — callers that need the name
   * should fall back to fetching the row themselves.
   */
  onRecordChange?: (id: string | null, primary?: string) => void;
  group?: RequestGroup;
  /** Description line under the title */
  sub?: string;
  /**
   * Reset every column-bound clause ($select / $filter / $orderby / $top /
   * $expand / $apply / ...) to its empty default, BUT keep the current
   * table. Surfaced as a "Reset request" button in the pane. Optional
   * because not every mode wants it (e.g. ExecuteAction has no concept of
   * a column-bound request).
   */
  onResetRequest?: () => void;
  /**
   * Apply a parsed OData URL to the mode's state. The mode receives the
   * ParsedRequest and decides which fields to use. Optional for the same
   * reason as onResetRequest.
   */
  onApplyParsed?: (parsed: ParsedRequest) => void;
}

export function TargetEditor({
  table,
  onTableChange,
  recordId,
  onRecordChange,
  group = 'read',
  sub = 'Pick the entity set this request operates on.',
  onResetRequest,
  onApplyParsed,
}: TargetEditorProps) {
  const { entities, loading: entitiesLoading, needsSetup, scopeMode } = useScopedEntities();
  const openSettings = useOpenSettings();
  const { loading: tableLoading } = useLiveTable(table || null);
  const tbl = findTable(table);
  const showRecord = onRecordChange !== undefined;

  // Pure Fluent v9 Combobox pattern from the official storybook:
  //   • controlled `value` for the typed input string
  //   • `freeform` so the user can type without picking an option
  //   • `clearable` for the native clear icon
  //   • manual filter on the options list (Fluent v9 doesn't auto-filter)
  // See https://storybooks.fluentui.dev/react/llms/components-combobox.txt
  // — "Combobox with filtering" + "Clearable Combobox" sections.
  const [value, setValue] = useState<string>(tbl?.displayName ?? '');

  // Sync external selection → input text (happens after the metadata fetch
  // resolves on initial load, or when a parent mode resets the table).
  useEffect(() => {
    setValue(tbl?.displayName ?? '');
  }, [tbl?.displayName]);

  const matching = useMemo(() => {
    const q = value.trim().toLowerCase();
    if (!q) return entities;
    return entities.filter(
      (t) =>
        t.displayName.toLowerCase().includes(q) ||
        t.logicalName.toLowerCase().includes(q) ||
        t.entitySetName.toLowerCase().includes(q),
    );
  }, [entities, value]);

  const onChange: ComboboxProps['onChange'] = (e) => {
    const next = (e.target as HTMLInputElement).value;
    setValue(next);
    // Clearing the input via the native `clearable` X also resets selection.
    if (next === '' && table) onTableChange('');
  };

  const onOptionSelect: ComboboxProps['onOptionSelect'] = (_, d) => {
    if (!d.optionValue) {
      // The clearable X dispatches a no-option select — propagate clear.
      onTableChange('');
      setValue('');
      return;
    }
    const sel = entities.find((t) => t.logicalName === d.optionValue);
    onTableChange(d.optionValue);
    setValue(sel?.displayName ?? '');
  };

  // Quick-action row — Refresh + Reset + Paste OData URL. Rendered only
  // when the hosting mode supplies the callbacks (every column-bound
  // retrieve mode does; specialty modes like ExecuteAction omit them).
  const hasActions = !!onResetRequest || !!onApplyParsed;

  // Refresh metadata: nukes every cached entity/attribute/relationship
  // record so the next read re-fetches with the current $select projection.
  // Useful when (a) the user changed a column in the maker, (b) DRS rolled
  // a new metadata-projection version (new flags appear after the cache
  // is cleared), or (c) suspected stale data.
  const onRefreshMetadata = () => {
    metadata.invalidateAll();
    // Re-warm the currently-selected entity so the user doesn't lose state.
    if (table) void metadata.getTable(table);
  };

  // Guidance banner — shown only after the entity list has finished loading
  // AND settings indicate the scope is unconfigured. The loading gate prevents
  // a flash for returning users whose persisted settings load async.
  const setupHint =
    scopeMode === 'publisher-solution'
      ? 'Open Settings → Query Scope and select a Publisher, then pick one or more Solutions to populate this list. Or switch the Entity Source to “All Entities”.'
      : 'Open Settings → Query Scope and select one or more Solutions to populate this list. Or switch the Entity Source to “All Entities”.';

  return (
    <div>
      {needsSetup && !entitiesLoading && (
        <MessageBar
          intent="warning"
          style={{ marginBottom: tokens.spacingVerticalM, width: '100%' }}
        >
          <MessageBarBody>
            <MessageBarTitle>Entity scope not configured</MessageBarTitle>
            {setupHint}
          </MessageBarBody>
          <MessageBarActions
            containerAction={
              <Button
                appearance="primary"
                size="small"
                icon={<Settings20Regular />}
                onClick={openSettings}
              >
                Open Settings
              </Button>
            }
          />
        </MessageBar>
      )}
      <PaneHead icon={Table20Filled} title="Target" sub={sub} group={group}>
        {hasActions && (
          <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
            {onApplyParsed && <PasteODataDialog onApply={onApplyParsed} />}
            <Button
              icon={<ArrowSync20Regular />}
              appearance="outline"
              size="small"
              onClick={onRefreshMetadata}
              title="Re-fetch every entity / attribute / relationship from Dataverse. Clears the in-memory metadata cache. Useful after editing the schema in the maker, or to pick up newly-added metadata flags."
            >
              Refresh metadata
            </Button>
            {onResetRequest && (
              <Button
                icon={<ArrowReset20Regular />}
                appearance="outline"
                size="small"
                onClick={onResetRequest}
                disabled={!table}
                title="Clear $select / $filter / $orderby / $top / $expand / $apply while keeping the current table"
              >
                Reset request
              </Button>
            )}
          </div>
        )}
      </PaneHead>

      <div style={{ display: 'flex', flexDirection: 'column', gap: 16, maxWidth: 560 }}>
        <Field
          label="Table"
          hint={
            tbl
              ? `${tbl.entitySetName} · ${tbl.columns.length} columns${tableLoading ? ' · loading…' : ''}`
              : entitiesLoading
                ? 'Loading entity list…'
                : `${entities.length} entities`
          }
        >
          <Combobox
            freeform
            clearable
            value={value}
            selectedOptions={table ? [table] : []}
            onChange={onChange}
            onOptionSelect={onOptionSelect}
            placeholder={entitiesLoading ? 'Loading…' : 'Search tables by name…'}
            listbox={{ style: { maxHeight: 360 } }}
          >
            {matching.length === 0 && (
              <Option value="__none" text="" disabled>
                <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                  No tables match "{value}"
                </Caption1>
              </Option>
            )}
            {matching.map((t) => (
              <Option key={t.logicalName} value={t.logicalName} text={t.displayName}>
                <div style={{ display: 'flex', flexDirection: 'column' }}>
                  <Body1>{t.displayName}</Body1>
                  <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                    {t.entitySetName} · {t.logicalName}
                  </Caption1>
                </div>
              </Option>
            ))}
          </Combobox>
        </Field>

        {tableLoading && (
          <div
            style={{
              display: 'flex',
              alignItems: 'center',
              gap: 8,
              color: tokens.colorNeutralForeground3,
            }}
          >
            <Spinner size="tiny" />
            <Caption1>Fetching columns + relationships from Dataverse…</Caption1>
          </div>
        )}

        {showRecord && tbl && (
          <Field label="Record">
            <RecordPicker
              table={table}
              selectedId={recordId ?? null}
              onPick={(r) => onRecordChange?.(r?.id ?? null, r?.primary ?? '')}
              placeholder={`Search ${tbl.displayName} records by ${tbl.primaryName}…`}
            />
          </Field>
        )}
      </div>
    </div>
  );
}
