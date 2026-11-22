// DisplaySettings — persisted, cross-mode preferences.
//
// Includes display flags (logical names, value-display mode) plus the
// active publisher / solution selections made with the header's scope chip.
// Persisted across reloads so the user's scope sticks.
//
// Results-grid column widths (#12) are stored here too, keyed per table —
// see views/results/columnWidths.ts.
//
// Storage namespace: `pptb-dataverse-request-studio:displaySettings` so the
// blob doesn't collide with other PPTB tools that persist their own
// preferences in localStorage.

import type { ColumnWidthStore } from '../views/results/columnWidths';

export type ValueDisplayMode = 'formatted' | 'raw' | 'both';

export type EntityScopeMode = 'publisher-solution' | 'solution-only' | 'all';

/** Results-grid row density — Fluent DataGrid's `size` values (#16). */
export type GridDensity = 'extra-small' | 'small' | 'medium';

export interface DisplaySettings {
  useLogicalNames: boolean;
  valueDisplayMode: ValueDisplayMode;
  /** Results-grid column widths the user dragged, per table + column (#12). */
  columnWidths: ColumnWidthStore;
  entityScopeMode: EntityScopeMode;
  advancedFindOnly: boolean;
  /** Results-grid row density (#16). */
  gridDensity: GridDensity;
  /**
   * Production safe mode (#241): on a Production connection every write asks
   * first, and you type the table (or operation) name to confirm.
   */
  productionSafeMode: boolean;
  /**
   * Selected publisher ids (publisherid GUIDs). Used when
   * entityScopeMode === 'publisher-solution' to narrow the solution list.
   * Empty array = no publisher chosen yet (the solutions list will be empty).
   */
  selectedPublisherIds: string[];
  /**
   * Selected solution ids (solutionid GUIDs). Used in both publisher-solution
   * and solution-only modes to narrow the entity list. Empty = no solutions
   * chosen yet (entity list will be empty in scoped modes).
   */
  selectedSolutionIds: string[];
  /**
   * The environments (org hosts) where the scope chip's first-run tip was
   * dismissed (#42). It shows once per environment, until a scope is set.
   */
  scopeTipDismissed: string[];
  /**
   * Keep each mode's runs across reloads (#226): the last 200 runs per
   * environment, without response bodies or credential / impersonation
   * headers. Off keeps Recent runs for this session only.
   */
  keepRunHistory: boolean;
}

/** Publisher / solution selection for one org — the GUIDs are per environment. */
export interface OrgScopeSelection {
  selectedPublisherIds: string[];
  selectedSolutionIds: string[];
}

/**
 * The persisted blob: the settings plus each org's own scope selection,
 * keyed by org host (#52). The top-level `selected*` fields hold the
 * selection last made in any org — older builds read those.
 */
export interface PersistedDisplaySettings extends DisplaySettings {
  scopeByOrg?: Record<string, OrgScopeSelection>;
  /**
   * 2 from 1.5.0. Before it, "Advanced Find only" did nothing and was saved
   * as `true` by default, so a value stored without this marker is dropped
   * once and the new default (off) applies (#44).
   */
  settingsRevision?: number;
}

/** The current `settingsRevision`. */
export const SETTINGS_REVISION = 2;

/**
 * The settings as seen from one org: that org's scope selection, or an
 * empty one if the org has none yet. Without an org (standalone, or no
 * connection yet) the top-level selection applies.
 */
export function settingsForOrg(stored: PersistedDisplaySettings, org: string): DisplaySettings {
  const { scopeByOrg, ...settings } = stored;
  if (!org || !scopeByOrg) return settings;
  const own = scopeByOrg[org];
  return {
    ...settings,
    selectedPublisherIds: own?.selectedPublisherIds ?? [],
    selectedSolutionIds: own?.selectedSolutionIds ?? [],
  };
}

/**
 * `stored` after the settings were changed to `next` while connected to
 * `org`: the shared fields, plus that org's scope selection (#52). Without
 * an org only the top-level selection changes.
 */
export function withSettingsChange(
  stored: PersistedDisplaySettings,
  next: DisplaySettings,
  org: string,
): PersistedDisplaySettings {
  const result: PersistedDisplaySettings = { ...stored, ...next };
  if (org) {
    result.scopeByOrg = {
      ...stored.scopeByOrg,
      [org]: {
        selectedPublisherIds: next.selectedPublisherIds,
        selectedSolutionIds: next.selectedSolutionIds,
      },
    };
  }
  return result;
}

/** The fields of `next` that differ from `prev` (compared by reference). */
export function changedSettings(
  prev: DisplaySettings,
  next: DisplaySettings,
): Partial<DisplaySettings> {
  const changed: Partial<DisplaySettings> = {};
  for (const key of Object.keys(next) as (keyof DisplaySettings)[]) {
    if (next[key] !== prev[key]) Object.assign(changed, { [key]: next[key] });
  }
  return changed;
}

export const defaultDisplaySettings: DisplaySettings = {
  useLogicalNames: false,
  valueDisplayMode: 'formatted',
  columnWidths: {},
  entityScopeMode: 'publisher-solution',
  // Off by default, as in FetchXML Studio (#44).
  advancedFindOnly: false,
  gridDensity: 'medium',
  productionSafeMode: true,
  selectedPublisherIds: [],
  selectedSolutionIds: [],
  scopeTipDismissed: [],
  keepRunHistory: true,
};

/** Settings namespace prefix. Exported for any future settings keys. */
export const SETTINGS_NAMESPACE = 'pptb-dataverse-request-studio';

/** Compose a namespaced settings key. */
export const settingsKey = (name: string): string => `${SETTINGS_NAMESPACE}:${name}`;

/** Storage key for the DisplaySettings blob. */
export const DISPLAY_SETTINGS_KEY = settingsKey('displaySettings');

/** Merge stored settings with defaults — graceful schema evolution. */
export function mergeWithDefaults(
  stored: Partial<PersistedDisplaySettings> | undefined,
): PersistedDisplaySettings {
  if (!stored) return { ...defaultDisplaySettings, settingsRevision: SETTINGS_REVISION };
  const migrated = { ...stored };
  if ((stored.settingsRevision ?? 1) < 2) delete migrated.advancedFindOnly;
  return { ...defaultDisplaySettings, ...migrated, settingsRevision: SETTINGS_REVISION };
}
