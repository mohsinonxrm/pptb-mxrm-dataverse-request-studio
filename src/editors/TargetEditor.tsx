import { useEffect, useState } from 'react';
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
  Filter20Regular,
} from '@fluentui/react-icons';
import { PaneHead } from './PaneHead';
import { findTable } from '../metadata/metadata';
import { useLiveTable } from '../host/useLiveMetadata';
import { useScopedEntities } from '../host/useScopedEntities';
import { useOpenScope } from '../host/useOpenScope';
import { SCOPE_TIP_BODY } from '../state/scopeChip';
import { RecordPicker } from '../primitives/RecordPicker';
import { MoreOptionsHint, useOptionSearch } from '../primitives/optionSearch';
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
  const {
    entities,
    loading: entitiesLoading,
    error: entitiesError,
    retry: retryEntities,
    needsSetup,
  } = useScopedEntities();
  const openScope = useOpenScope();
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

  // Capped + deferred filter: an org can have ~2,000 tables and Fluent
  // renders every Option on focus (#21).
  const {
    visible: matching,
    matchCount,
    hiddenCount,
  } = useOptionSearch(entities, value, (t) => [t.displayName, t.logicalName, t.entitySetName]);

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

  // Refresh metadata: drops every cached entity/attribute/relationship
  // record and rebuilds the loaded tables in place, same as Settings →
  // Refresh metadata. Useful when (a) the user changed a column in the
  // maker, (b) DRS rolled a new metadata-projection version (new flags
  // appear after the cache is cleared), or (c) suspected stale data.
  // In place, so the selected table never blanks out mid-refresh (#53).
  const [refreshing, setRefreshing] = useState(false);
  const onRefreshMetadata = () => {
    setRefreshing(true);
    metadata
      .refreshAll()
      .catch((e) => console.warn('[TargetEditor] refresh metadata failed', e))
      .finally(() => setRefreshing(false));
  };

  // No scope yet (#42): the shared empty state, and the picker disabled.
  // Only after the entity list has finished loading — the loading gate
  // prevents a flash for returning users whose persisted settings load async.
  const noScope = needsSetup && !entitiesLoading;

  return (
    <div>
      {noScope && (
        <MessageBar intent="info" style={{ marginBottom: tokens.spacingVerticalM, width: '100%' }}>
          <MessageBarBody>
            <MessageBarTitle>No entities listed yet</MessageBarTitle>
            {SCOPE_TIP_BODY}
          </MessageBarBody>
          <MessageBarActions
            containerAction={
              <Button
                appearance="primary"
                size="small"
                icon={<Filter20Regular />}
                onClick={openScope}
              >
                Choose scope
              </Button>
            }
          />
        </MessageBar>
      )}
      {entitiesError && !entitiesLoading && (
        <MessageBar intent="error" style={{ marginBottom: tokens.spacingVerticalM, width: '100%' }}>
          <MessageBarBody>
            <MessageBarTitle>Couldn't load the table list</MessageBarTitle>
            {entitiesError}
          </MessageBarBody>
          <MessageBarActions
            containerAction={
              <Button
                appearance="primary"
                size="small"
                icon={<ArrowSync20Regular />}
                onClick={retryEntities}
              >
                Retry
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
              disabled={refreshing}
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
                : entitiesError
                  ? 'Entity list failed to load'
                  : `${entities.length} ${entities.length === 1 ? 'entity' : 'entities'} in scope`
          }
        >
          <Combobox
            freeform
            clearable
            value={value}
            selectedOptions={table ? [table] : []}
            onChange={onChange}
            onOptionSelect={onOptionSelect}
            // A table already picked (a restored request, another org) can
            // still be cleared.
            disabled={noScope && !table}
            placeholder={
              entitiesLoading
                ? 'Loading…'
                : noScope
                  ? 'Choose a scope to list entities'
                  : 'Search tables by name…'
            }
            listbox={{ style: { maxHeight: 360 } }}
          >
            {matchCount === 0 && (
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
            <MoreOptionsHint hiddenCount={hiddenCount} />
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
              memoryKey="target"
            />
          </Field>
        )}
      </div>
    </div>
  );
}
