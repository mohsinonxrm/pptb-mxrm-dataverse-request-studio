// DisplaySettings — persisted, cross-mode preferences.
//
// Includes display flags (logical names, value-display mode) plus the
// active publisher / solution selections sourced from the Settings drawer.
// Persisted across reloads so the user's scope sticks.
//
// Storage namespace: `pptb-dataverse-request-studio:displaySettings` so the
// blob doesn't collide with other PPTB tools that persist their own
// preferences in localStorage.

export type ValueDisplayMode = 'formatted' | 'raw' | 'both';

export type EntityScopeMode = 'publisher-solution' | 'solution-only' | 'all';

export interface DisplaySettings {
  useLogicalNames: boolean;
  valueDisplayMode: ValueDisplayMode;
  entityScopeMode: EntityScopeMode;
  advancedFindOnly: boolean;
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
  entityScopeMode: 'publisher-solution',
  advancedFindOnly: true,
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
export function mergeWithDefaults(
  stored: Partial<DisplaySettings> | undefined,
): DisplaySettings {
  if (!stored) return { ...defaultDisplaySettings };
  return { ...defaultDisplaySettings, ...stored };
}
