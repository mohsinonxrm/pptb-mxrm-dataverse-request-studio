// SettingsDrawer — Fluent v9 OverlayDrawer hosting query-scope + display
// preferences.
//
// Three sections:
//
//   1. Query Scope
//      - Entity Source RadioGroup (publisher-solution / solution-only / all)
//        gated by AccessSummary (privilege check)
//      - Publisher multi-pick (TagPicker)         [publisher-solution only]
//      - Solution multi-pick (TagPicker)          [publisher-solution + solution-only]
//      - Advanced Find Only Switch
//
//   2. Display
//      - Use Logical Names Switch
//      - Value Display Mode Dropdown (Formatted / Raw / Both)
//
// Publisher / solution pickers live here in the drawer (rather than in the
// toolbar). Selections are persisted on DisplaySettings so they survive
// reloads.

import { useCallback, useEffect, useMemo, useState } from 'react';
import {
  DrawerBody, DrawerHeader, DrawerHeaderTitle, OverlayDrawer,
  Button, makeStyles, tokens, Text, Switch, Dropdown, Option,
  Divider, Radio, RadioGroup, Tooltip, Caption1, Spinner,
  TagPicker, TagPickerControl, TagPickerGroup, TagPickerInput,
  TagPickerList, TagPickerOption, useTagPickerFilter, Tag,
} from '@fluentui/react-components';
import { Dismiss24Regular, Settings20Regular, Info16Regular, ArrowClockwise20Regular } from '@fluentui/react-icons';
import type {
  DisplaySettings, ValueDisplayMode, EntityScopeMode,
} from '../state/displaySettings';
import type { AccessSummary } from '../host/pptbClient';
import { metadata } from '../host/metadataProvider';
import { usePublisherFilter } from '../host/usePublisherFilter';
import { useSolutionFilter } from '../host/useSolutionFilter';

const useStyles = makeStyles({
  drawer: { width: '360px' },
  section: { marginBottom: tokens.spacingVerticalL },
  sectionTitle: {
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
    marginBottom: tokens.spacingVerticalM,
    display: 'flex',
    alignItems: 'center',
    columnGap: tokens.spacingHorizontalS,
  },
  settingItem: {
    display: 'flex',
    flexDirection: 'column',
    rowGap: tokens.spacingVerticalXS,
    marginBottom: tokens.spacingVerticalM,
  },
  settingRow: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
  },
  settingLabel: {
    fontWeight: tokens.fontWeightSemibold,
    display: 'flex',
    alignItems: 'center',
    columnGap: tokens.spacingHorizontalXS,
  },
  settingDescription: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
  },
  dropdown: { minWidth: '160px' },
  radioGroup: {
    display: 'flex',
    flexDirection: 'column',
    rowGap: tokens.spacingVerticalXS,
    marginTop: tokens.spacingVerticalXS,
  },
  disabledHint: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForegroundDisabled,
    fontStyle: 'italic',
  },
});

export interface SettingsDrawerProps {
  open: boolean;
  settings: DisplaySettings;
  onClose: () => void;
  onSettingsChange: (settings: DisplaySettings) => void;
  accessSummary?: AccessSummary | null;
}

export function SettingsDrawer({
  open, settings, onClose, onSettingsChange, accessSummary,
}: SettingsDrawerProps) {
  const styles = useStyles();

  const handleLogicalNamesChange = useCallback((checked: boolean) => {
    onSettingsChange({ ...settings, useLogicalNames: checked });
  }, [settings, onSettingsChange]);

  const handleValueDisplayModeChange = useCallback((mode: ValueDisplayMode) => {
    onSettingsChange({ ...settings, valueDisplayMode: mode });
  }, [settings, onSettingsChange]);

  const handleEntityScopeModeChange = useCallback((mode: EntityScopeMode) => {
    onSettingsChange({ ...settings, entityScopeMode: mode });
  }, [settings, onSettingsChange]);

  const handleAdvancedFindOnlyChange = useCallback((checked: boolean) => {
    onSettingsChange({ ...settings, advancedFindOnly: checked });
  }, [settings, onSettingsChange]);

  const handlePublisherIdsChange = useCallback((ids: string[]) => {
    onSettingsChange({ ...settings, selectedPublisherIds: ids });
  }, [settings, onSettingsChange]);

  const handleSolutionIdsChange = useCallback((ids: string[]) => {
    onSettingsChange({ ...settings, selectedSolutionIds: ids });
  }, [settings, onSettingsChange]);

  // Refresh metadata — drops every cache entry and re-fetches the
  // currently-loaded tables in place. Used when entity/attribute/relationship
  // definitions change in the environment mid-session (the cache TTL is 1h).
  const [refreshing, setRefreshing] = useState(false);
  const handleRefreshMetadata = useCallback(async () => {
    setRefreshing(true);
    try { await metadata.refreshAll(); }
    catch (e) { console.warn('[SettingsDrawer] refresh metadata failed', e); }
    finally { setRefreshing(false); }
  }, []);

  // Privilege gating — when AccessSummary hasn't loaded yet we allow all
  // options (deferred validation).
  const canUsePublisherSolution = !accessSummary || accessSummary.fullFilterMode;
  const canUseSolutionOnly =
    !accessSummary || accessSummary.fullFilterMode || accessSummary.solutionsOnlyMode;

  return (
    <OverlayDrawer
      open={open}
      onOpenChange={(_e, data) => { if (!data.open) onClose(); }}
      position="end"
      size="small"
      className={styles.drawer}
    >
      <DrawerHeader>
        <DrawerHeaderTitle
          action={
            <Button
              appearance="subtle"
              aria-label="Close"
              icon={<Dismiss24Regular />}
              onClick={onClose}
            />
          }
        >
          <Settings20Regular style={{ marginRight: tokens.spacingHorizontalS }} />
          Settings
        </DrawerHeaderTitle>
      </DrawerHeader>
      <DrawerBody>
        {/* ── Section: Query Scope ─────────────────────────────── */}
        <div className={styles.section}>
          <Text className={styles.sectionTitle}>Query Scope</Text>

          <div className={styles.settingItem}>
            <Text className={styles.settingLabel}>
              Entity Source
              <Tooltip
                content="Controls how the entity list is populated in the toolbar. Scoped modes are faster; 'All Entities' loads every entity in the environment."
                relationship="description"
              >
                <Info16Regular style={{ color: tokens.colorNeutralForeground3 }} />
              </Tooltip>
            </Text>
            <Text className={styles.settingDescription} block>
              How to scope the entity picker
            </Text>
            <RadioGroup
              className={styles.radioGroup}
              value={settings.entityScopeMode}
              onChange={(_e, data) => handleEntityScopeModeChange(data.value as EntityScopeMode)}
            >
              <Tooltip
                content={!canUsePublisherSolution ? 'Requires prvReadPublisher + prvReadSolution' : ''}
                relationship="description"
                positioning="before"
              >
                <Radio
                  value="publisher-solution"
                  label="Publisher → Solution (default)"
                  disabled={!canUsePublisherSolution}
                />
              </Tooltip>
              <Tooltip
                content={!canUseSolutionOnly ? 'Requires prvReadSolution' : ''}
                relationship="description"
                positioning="before"
              >
                <Radio
                  value="solution-only"
                  label="Solution only"
                  disabled={!canUseSolutionOnly}
                />
              </Tooltip>
              <Radio value="all" label="All Entities" />
            </RadioGroup>
          </div>

          {/* Publisher + Solution pickers — only meaningful in scoped modes. */}
          {settings.entityScopeMode === 'publisher-solution' && (
            <PublisherSolutionPickers
              selectedPublisherIds={settings.selectedPublisherIds}
              selectedSolutionIds={settings.selectedSolutionIds}
              onPublisherIdsChange={handlePublisherIdsChange}
              onSolutionIdsChange={handleSolutionIdsChange}
            />
          )}
          {settings.entityScopeMode === 'solution-only' && (
            <SolutionOnlyPicker
              selectedSolutionIds={settings.selectedSolutionIds}
              onSolutionIdsChange={handleSolutionIdsChange}
            />
          )}

          <Divider style={{ marginBottom: tokens.spacingVerticalM }} />

          <div className={styles.settingItem}>
            <div className={styles.settingRow}>
              <div>
                <Text className={styles.settingLabel}>
                  Advanced Find Only
                  <Tooltip
                    content="When on, entities and attributes are limited to those marked IsValidForAdvancedFind — the same set exposed by Advanced Find in Model-Driven Apps. Turn off to access all entities and attributes."
                    relationship="description"
                  >
                    <Info16Regular style={{ color: tokens.colorNeutralForeground3 }} />
                  </Tooltip>
                </Text>
                <Text className={styles.settingDescription} block>
                  Limit entities &amp; attributes to Advanced Find-eligible ones
                </Text>
                {!settings.advancedFindOnly && (
                  <Text className={styles.disabledHint} block>
                    All entities and attributes are shown — including non-queryable ones
                  </Text>
                )}
              </div>
              <Switch
                checked={settings.advancedFindOnly}
                onChange={(_e, data) => handleAdvancedFindOnlyChange(data.checked)}
              />
            </div>
          </div>
        </div>

        <Divider />

        {/* ── Section: Display ─────────────────────────────────── */}
        <div className={styles.section} style={{ marginTop: tokens.spacingVerticalL }}>
          <Text className={styles.sectionTitle}>Display</Text>

          <div className={styles.settingItem}>
            <div className={styles.settingRow}>
              <div>
                <Text className={styles.settingLabel}>Use Logical Names</Text>
                <Text className={styles.settingDescription} block>
                  Show attribute logical names in column headers instead of display names
                </Text>
              </div>
              <Switch
                checked={settings.useLogicalNames}
                onChange={(_e, data) => handleLogicalNamesChange(data.checked)}
              />
            </div>
          </div>

          <div className={styles.settingItem}>
            <Text className={styles.settingLabel}>Value Display Mode</Text>
            <Text className={styles.settingDescription} block>
              How to display cell values in the results grid
            </Text>
            <Dropdown
              className={styles.dropdown}
              value={
                settings.valueDisplayMode === 'formatted' ? 'Formatted' :
                settings.valueDisplayMode === 'raw' ? 'Raw' :
                'Both (2 columns per attribute)'
              }
              selectedOptions={[settings.valueDisplayMode]}
              onOptionSelect={(_e, data) =>
                handleValueDisplayModeChange(data.optionValue as ValueDisplayMode)
              }
            >
              <Option value="formatted">Formatted</Option>
              <Option value="raw">Raw</Option>
              <Option value="both">Both (2 columns per attribute)</Option>
            </Dropdown>
          </div>
        </div>

        <Divider />

        {/* ── Section: Metadata ────────────────────────────────── */}
        <div className={styles.section} style={{ marginTop: tokens.spacingVerticalL }}>
          <Text className={styles.sectionTitle}>Metadata</Text>

          <div className={styles.settingItem}>
            <Text className={styles.settingLabel}>Refresh metadata</Text>
            <Text className={styles.settingDescription} block style={{ marginBottom: tokens.spacingVerticalS }}>
              Re-fetch entity, attribute &amp; relationship definitions from the
              environment. Use this after publishing schema changes — metadata is
              otherwise cached for up to an hour.
            </Text>
            <Button
              appearance="outline"
              size="small"
              icon={refreshing ? <Spinner size="tiny" /> : <ArrowClockwise20Regular />}
              disabled={refreshing}
              onClick={handleRefreshMetadata}
              style={{ alignSelf: 'flex-start' }}
            >
              {refreshing ? 'Refreshing…' : 'Refresh metadata'}
            </Button>
          </div>
        </div>
      </DrawerBody>
    </OverlayDrawer>
  );
}

// ───────────────────────────────────────────────────────────────────────
// Publisher → Solution cascade picker (entityScopeMode === 'publisher-solution')
// ───────────────────────────────────────────────────────────────────────
function PublisherSolutionPickers({
  selectedPublisherIds, selectedSolutionIds,
  onPublisherIdsChange, onSolutionIdsChange,
}: {
  selectedPublisherIds: string[];
  selectedSolutionIds: string[];
  onPublisherIdsChange: (ids: string[]) => void;
  onSolutionIdsChange: (ids: string[]) => void;
}) {
  // The hook owns its own internal selection state. We sync settings →
  // hook on mount and hook → settings on any user change below.
  const filter = usePublisherFilter();
  const {
    publishers, publishersLoading, publishersError,
    solutions, updateSelectedPublishers,
  } = filter;

  // Sync persisted selections into the hook on mount + when settings change.
  useEffect(() => {
    updateSelectedPublishers(selectedPublisherIds);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedPublisherIds.join('|')]);

  const publisherById = useMemo(() => {
    const m = new Map<string, string>();
    for (const p of publishers) m.set(p.publisherid, p.friendlyname || p.uniquename);
    return m;
  }, [publishers]);
  const solutionById = useMemo(() => {
    const m = new Map<string, string>();
    for (const s of solutions) m.set(s.solutionid, s.friendlyname || s.uniquename);
    return m;
  }, [solutions]);

  return (
    <>
      <div style={{ marginTop: tokens.spacingVerticalM, marginBottom: tokens.spacingVerticalM }}>
        <Text style={{ fontWeight: tokens.fontWeightSemibold, display: 'block', marginBottom: 4 }}>
          Publisher
        </Text>
        <Caption1 style={{ color: tokens.colorNeutralForeground3, display: 'block', marginBottom: 6 }}>
          Pick one or more publishers — only their solutions appear below.
        </Caption1>
        {publishersLoading && (
          <div style={{ display: 'flex', gap: 8, alignItems: 'center', padding: '4px 0' }}>
            <Spinner size="extra-small" />
            <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>Loading publishers…</Caption1>
          </div>
        )}
        {publishersError && (
          <Caption1 style={{ color: tokens.colorPaletteRedForeground1 }}>
            {publishersError}
          </Caption1>
        )}
        <MultiPicker
          values={selectedPublisherIds}
          options={publishers.map(p => p.publisherid)}
          labelByKey={publisherById}
          placeholder="Search publishers…"
          onChange={onPublisherIdsChange}
        />
      </div>

      <div style={{ marginBottom: tokens.spacingVerticalM }}>
        <Text style={{ fontWeight: tokens.fontWeightSemibold, display: 'block', marginBottom: 4 }}>
          Solution
        </Text>
        <Caption1 style={{ color: tokens.colorNeutralForeground3, display: 'block', marginBottom: 6 }}>
          {selectedPublisherIds.length === 0
            ? 'Pick a publisher first to populate this list.'
            : `${solutions.length} solution${solutions.length === 1 ? '' : 's'} from the selected publisher${selectedPublisherIds.length === 1 ? '' : 's'}.`}
        </Caption1>
        <MultiPicker
          values={selectedSolutionIds}
          options={solutions.map(s => s.solutionid)}
          labelByKey={solutionById}
          placeholder="Search solutions…"
          onChange={onSolutionIdsChange}
          disabled={selectedPublisherIds.length === 0}
        />
      </div>
    </>
  );
}

// ───────────────────────────────────────────────────────────────────────
// Solution-only picker (entityScopeMode === 'solution-only')
// ───────────────────────────────────────────────────────────────────────
function SolutionOnlyPicker({
  selectedSolutionIds, onSolutionIdsChange,
}: {
  selectedSolutionIds: string[];
  onSolutionIdsChange: (ids: string[]) => void;
}) {
  const { solutions, solutionsLoading, solutionsError } = useSolutionFilter();
  const solutionById = useMemo(() => {
    const m = new Map<string, string>();
    for (const s of solutions) m.set(s.solutionid, s.friendlyname || s.uniquename);
    return m;
  }, [solutions]);

  return (
    <div style={{ marginTop: tokens.spacingVerticalM, marginBottom: tokens.spacingVerticalM }}>
      <Text style={{ fontWeight: tokens.fontWeightSemibold, display: 'block', marginBottom: 4 }}>
        Solution
      </Text>
      <Caption1 style={{ color: tokens.colorNeutralForeground3, display: 'block', marginBottom: 6 }}>
        Pick one or more solutions to scope the entity list to their components.
      </Caption1>
      {solutionsLoading && (
        <div style={{ display: 'flex', gap: 8, alignItems: 'center', padding: '4px 0' }}>
          <Spinner size="extra-small" />
          <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>Loading solutions…</Caption1>
        </div>
      )}
      {solutionsError && (
        <Caption1 style={{ color: tokens.colorPaletteRedForeground1 }}>
          {solutionsError}
        </Caption1>
      )}
      <MultiPicker
        values={selectedSolutionIds}
        options={solutions.map(s => s.solutionid)}
        labelByKey={solutionById}
        placeholder="Search solutions…"
        onChange={onSolutionIdsChange}
      />
    </div>
  );
}

// ───────────────────────────────────────────────────────────────────────
// Reusable Fluent v9 TagPicker (filterable multi-select over string keys)
// ───────────────────────────────────────────────────────────────────────
function MultiPicker({
  values, options, labelByKey, placeholder, disabled, onChange,
}: {
  values: string[];
  options: string[];
  labelByKey: Map<string, string>;
  placeholder?: string;
  disabled?: boolean;
  onChange: (next: string[]) => void;
}) {
  const [query, setQuery] = useState('');

  const children = useTagPickerFilter({
    query,
    options,
    noOptionsElement: <TagPickerOption value="__no-matches">No matches</TagPickerOption>,
    renderOption: (key) => (
      <TagPickerOption key={key} value={key} text={labelByKey.get(key) ?? key}>
        {labelByKey.get(key) ?? key}
      </TagPickerOption>
    ),
    filter: (key) => {
      if (values.includes(key)) return false;
      const label = labelByKey.get(key) ?? key;
      return label.toLowerCase().includes(query.toLowerCase());
    },
  });

  return (
    <TagPicker
      disabled={disabled}
      selectedOptions={values}
      onOptionSelect={(_e, data) => {
        if (data.value === '__no-matches') return;
        onChange(data.selectedOptions);
        setQuery('');
      }}
    >
      <TagPickerControl style={{ minWidth: 0 }}>
        <TagPickerGroup aria-label="Selected">
          {values.map(k => (
            <Tag key={k} value={k} shape="rounded" size="small"
              dismissible
              dismissIcon={{ 'aria-label': 'remove' }}
            >
              {labelByKey.get(k) ?? k}
            </Tag>
          ))}
        </TagPickerGroup>
        <TagPickerInput
          aria-label="Pick option"
          placeholder={values.length === 0 ? placeholder : ''}
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          disabled={disabled}
        />
      </TagPickerControl>
      <TagPickerList>{children}</TagPickerList>
    </TagPicker>
  );
}
