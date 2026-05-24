// SettingsProvider + usePersistedSettings — shared, context-backed
// DisplaySettings with toolboxAPI persistence.
//
// Previously each call to `usePersistedSettings` created its own React
// state, so an update in the Settings drawer didn't reach
// `useScopedEntities` (different consumer, different state). Now the
// provider at the top of the tree owns the single source of truth, and
// every consumer reads/writes the same store. Persistence still flows
// through `window.toolboxAPI.settings`.
//
// `loadedRef` gates the write-back so the initial read doesn't trigger
// a save before we've actually read what's in storage. Standalone mode
// (no toolboxAPI): settings stay in memory and reset on reload.

import {
  createContext, useCallback, useContext, useEffect, useRef, useState,
  type ReactNode,
} from 'react';
import {
  defaultDisplaySettings, mergeWithDefaults, DISPLAY_SETTINGS_KEY,
  type DisplaySettings,
} from '../state/displaySettings';

type Ctx = [DisplaySettings, (next: DisplaySettings) => void];

const SettingsContext = createContext<Ctx | null>(null);

export function SettingsProvider({ children }: { children: ReactNode }) {
  const [settings, setSettingsState] = useState<DisplaySettings>(defaultDisplaySettings);
  const loadedRef = useRef(false);

  // ── Load on mount ──
  useEffect(() => {
    let mounted = true;

    async function loadSettings() {
      try {
        if (typeof window !== 'undefined' && window.toolboxAPI?.settings?.get) {
          const persisted = await window.toolboxAPI.settings.get(DISPLAY_SETTINGS_KEY);
          if (mounted) {
            setSettingsState(mergeWithDefaults(persisted as Partial<DisplaySettings> | undefined));
          }
        }
      } catch (err) {
        console.warn('⚙️ Failed to load settings from toolboxAPI:', err);
      } finally {
        if (mounted) loadedRef.current = true;
      }
    }

    loadSettings();
    return () => { mounted = false; };
  }, []);

  // ── Update + persist ──
  const updateSettings = useCallback((newSettings: DisplaySettings) => {
    setSettingsState(newSettings);

    if (loadedRef.current) {
      if (typeof window !== 'undefined' && window.toolboxAPI?.settings?.set) {
        window.toolboxAPI.settings
          .set(DISPLAY_SETTINGS_KEY, newSettings)
          .catch((err: unknown) => {
            console.warn('⚙️ Failed to save settings to toolboxAPI:', err);
          });
      }
    }
  }, []);

  return (
    <SettingsContext.Provider value={[settings, updateSettings]}>
      {children}
    </SettingsContext.Provider>
  );
}

/**
 * Returns `[settings, updateSettings]` from the SettingsProvider context.
 * Falls back to a no-op pair if accidentally called outside the provider
 * so renders don't crash — but warn loudly so it gets noticed.
 */
export function usePersistedSettings(): Ctx {
  const ctx = useContext(SettingsContext);
  if (!ctx) {
    if (typeof window !== 'undefined') {
      console.warn(
        '[usePersistedSettings] called outside <SettingsProvider> — ' +
        'wrap the app in <SettingsProvider> in App.tsx.',
      );
    }
    return [defaultDisplaySettings, () => {}];
  }
  return ctx;
}
