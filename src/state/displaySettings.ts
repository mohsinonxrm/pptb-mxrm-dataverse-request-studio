// DisplaySettings — persisted, cross-mode preferences.
//
// Includes display flags (logical names, value-display mode) plus the
// active publisher / solution selections sourced from the Settings drawer.
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
}

export const defaultDisplaySettings: DisplaySettings = {
  useLogicalNames: false,
  valueDisplayMode: 'formatted',
  columnWidths: {},
  entityScopeMode: 'publisher-solution',
  advancedFindOnly: true,
  gridDensity: 'medium',
  selectedPublisherIds: [],
  selectedSolutionIds: [],
};

/** Settings namespace prefix. Exported for any future settings keys. */
export const SETTINGS_NAMESPACE = 'pptb-dataverse-request-studio';

/** Compose a namespaced settings key. */
export const settingsKey = (name: string): string => `${SETTINGS_NAMESPACE}:${name}`;

/** Storage key for the DisplaySettings blob. */
export const DISPLAY_SETTINGS_KEY = settingsKey('displaySettings');

/** Merge stored settings with defaults — graceful schema evolution. */
export function mergeWithDefaults(stored: Partial<DisplaySettings> | undefined): DisplaySettings {
  if (!stored) return { ...defaultDisplaySettings };
  return { ...defaultDisplaySettings, ...stored };
}
