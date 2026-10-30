// SettingsDrawer — Fluent v9 OverlayDrawer hosting display, safety and
// metadata preferences.
//
// Sections: Entity scope (a pointer to the header's scope chip, which owns
// the publisher / solution pickers since #42), Display, Safety, Metadata.
// Everything is persisted on DisplaySettings so it survives reloads.

import { useCallback, useState } from 'react';
import {
  DrawerBody,
  DrawerHeader,
  DrawerHeaderTitle,
  OverlayDrawer,
  Button,
  makeStyles,
  tokens,
  Text,
  Switch,
  Dropdown,
  Option,
  Divider,
  Spinner,
} from '@fluentui/react-components';
import {
  Dismiss24Regular,
  Settings20Regular,
  ArrowClockwise20Regular,
} from '@fluentui/react-icons';
import type { DisplaySettings, ValueDisplayMode } from '../state/displaySettings';
import { metadata } from '../host/metadataProvider';

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
});

export interface SettingsDrawerProps {
  open: boolean;
  settings: DisplaySettings;
  onClose: () => void;
  onSettingsChange: (settings: DisplaySettings) => void;
}

export function SettingsDrawer({ open, settings, onClose, onSettingsChange }: SettingsDrawerProps) {
  const styles = useStyles();

  const handleLogicalNamesChange = useCallback(
    (checked: boolean) => {
      onSettingsChange({ ...settings, useLogicalNames: checked });
    },
    [settings, onSettingsChange],
  );

  const handleSafeModeChange = useCallback(
    (checked: boolean) => {
      onSettingsChange({ ...settings, productionSafeMode: checked });
    },
    [settings, onSettingsChange],
  );

  const handleValueDisplayModeChange = useCallback(
    (mode: ValueDisplayMode) => {
      onSettingsChange({ ...settings, valueDisplayMode: mode });
    },
    [settings, onSettingsChange],
  );

  // Refresh metadata — drops every cache entry and re-fetches the
  // currently-loaded tables in place. Used when entity/attribute/relationship
  // definitions change in the environment mid-session (the cache TTL is 1h).
  const [refreshing, setRefreshing] = useState(false);
  const handleRefreshMetadata = useCallback(async () => {
    setRefreshing(true);
    try {
      await metadata.refreshAll();
    } catch (e) {
      console.warn('[SettingsDrawer] refresh metadata failed', e);
    } finally {
      setRefreshing(false);
    }
  }, []);

  return (
    <OverlayDrawer
      open={open}
      onOpenChange={(_e, data) => {
        if (!data.open) onClose();
      }}
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
        {/* ── Section: Entity scope (#42) ───────────────────────── */}
        <div className={styles.section}>
          <Text className={styles.sectionTitle}>Entity scope</Text>
          <Text className={styles.settingDescription} block>
            Set the entity scope from the scope button.
          </Text>
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
                aria-label="Use Logical Names"
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
              aria-label="Value Display Mode"
              className={styles.dropdown}
              value={
                settings.valueDisplayMode === 'formatted'
                  ? 'Formatted'
                  : settings.valueDisplayMode === 'raw'
                    ? 'Raw'
                    : 'Both (2 columns per attribute)'
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

        {/* ── Section: Safety (#241) ───────────────────────────── */}
        <div className={styles.section} style={{ marginTop: tokens.spacingVerticalL }}>
          <Text className={styles.sectionTitle}>Safety</Text>

          <div className={styles.settingItem}>
            <div className={styles.settingRow}>
              <div>
                <Text className={styles.settingLabel}>Production safe mode</Text>
                <Text className={styles.settingDescription} block>
                  On a Production connection, every write asks first, and you type the table (or
                  operation) name to confirm. Merge, actions, workflows, bulk and batch writes
                  always ask, on any connection.
                </Text>
              </div>
              <Switch
                aria-label="Production safe mode"
                checked={settings.productionSafeMode}
                onChange={(_e, data) => handleSafeModeChange(data.checked)}
              />
            </div>
          </div>
        </div>

        <Divider />

        {/* ── Section: Metadata ────────────────────────────────── */}
        <div className={styles.section} style={{ marginTop: tokens.spacingVerticalL }}>
          <Text className={styles.sectionTitle}>Metadata</Text>

          <div className={styles.settingItem}>
            <Text className={styles.settingLabel}>Refresh metadata</Text>
            <Text
              className={styles.settingDescription}
              block
              style={{ marginBottom: tokens.spacingVerticalS }}
            >
              Re-fetch entity, attribute &amp; relationship definitions from the environment. Use
              this after publishing schema changes — metadata is otherwise cached for up to an hour.
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
