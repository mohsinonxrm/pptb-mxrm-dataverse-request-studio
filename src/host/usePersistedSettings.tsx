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
// `loadedRef` gates the write-back so nothing is saved before we've
// actually read what's in storage. A failed read leaves it closed (the next
// change retries), and changes made before the read are merged over what
// it returns (#103). Standalone mode (no toolboxAPI): settings stay in
// memory and reset on reload.
//
// The publisher / solution selection is stored per org (#52) — its GUIDs
// only exist in the environment they were picked in. Consumers see the
// active org's selection; an org without one starts unconfigured, which
// shows the "Entity scope not configured" banner instead of an empty list.
// The first org seen after upgrading inherits the old global selection.

import {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useMemo,
  useRef,
  useState,
  type ReactNode,
} from 'react';
import {
  defaultDisplaySettings,
  changedSettings,
  mergeWithDefaults,
  settingsForOrg,
  withSettingsChange,
  DISPLAY_SETTINGS_KEY,
  type DisplaySettings,
  type PersistedDisplaySettings,
} from '../state/displaySettings';
import { useActiveOrg } from './activeOrg';

type Ctx = [DisplaySettings, (next: DisplaySettings) => void];

const SettingsContext = createContext<Ctx | null>(null);

function persist(next: PersistedDisplaySettings): void {
  if (typeof window !== 'undefined' && window.toolboxAPI?.settings?.set) {
    window.toolboxAPI.settings.set(DISPLAY_SETTINGS_KEY, next).catch((err: unknown) => {
      console.warn('⚙️ Failed to save settings to toolboxAPI:', err);
    });
  }
}

export function SettingsProvider({ children }: { children: ReactNode }) {
  const [stored, setStored] = useState<PersistedDisplaySettings>(defaultDisplaySettings);
  const storedRef = useRef(stored);
  storedRef.current = stored;
  const loadedRef = useRef(false);
  const org = useActiveOrg();
  const orgRef = useRef(org);
  orgRef.current = org;
  const settings = useMemo(() => settingsForOrg(stored, org), [stored, org]);
  // Changes made before the stored settings loaded — only the fields each
  // change touched, so they can be laid over what's in storage (#103).
  const pendingRef = useRef<{ org: string; changed: Partial<DisplaySettings> }[]>([]);
  const loadingRef = useRef<Promise<void> | null>(null);
  const mountedRef = useRef(true);

  // ── Load ──
  // `loadedRef` only turns true once storage was actually read. Until then
  // nothing is written back, so a failed read can't replace the stored
  // settings with the defaults (#103); the next change retries the read.
  const loadSettings = useCallback((): Promise<void> => {
    if (loadingRef.current) return loadingRef.current;
    const run = (async () => {
      try {
        if (typeof window === 'undefined' || !window.toolboxAPI?.settings?.get) {
          loadedRef.current = true;
          return;
        }
        const persisted = await window.toolboxAPI.settings.get(DISPLAY_SETTINGS_KEY);
        if (!mountedRef.current) return;
        const pending = pendingRef.current;
        pendingRef.current = [];
        let next = mergeWithDefaults(persisted as Partial<PersistedDisplaySettings> | undefined);
        for (const { org: changeOrg, changed } of pending) {
          next = withSettingsChange(
            next,
            { ...settingsForOrg(next, changeOrg), ...changed },
            changeOrg,
          );
        }
        loadedRef.current = true;
        storedRef.current = next;
        setStored(next);
        if (pending.length > 0) persist(next);
      } catch (err) {
        console.warn('⚙️ Failed to load settings from toolboxAPI:', err);
      } finally {
        loadingRef.current = null;
      }
    })();
    loadingRef.current = run;
    return run;
  }, []);

  useEffect(() => {
    mountedRef.current = true;
    void loadSettings();
    return () => {
      mountedRef.current = false;
    };
  }, [loadSettings]);

  // ── One-time migration ──
  // Settings saved before per-org scopes have only the global selection.
  // Pin it to the first org we see, so other orgs start unconfigured.
  useEffect(() => {
    if (!loadedRef.current || !org || stored.scopeByOrg) return;
    const next: PersistedDisplaySettings = {
      ...stored,
      scopeByOrg: {
        [org]: {
          selectedPublisherIds: stored.selectedPublisherIds,
          selectedSolutionIds: stored.selectedSolutionIds,
        },
      },
    };
    setStored(next);
    persist(next);
  }, [org, stored]);

  // ── Update + persist ──
  const updateSettings = useCallback(
    (newSettings: DisplaySettings) => {
      const currentOrg = orgRef.current;
      if (!loadedRef.current) {
        pendingRef.current.push({
          org: currentOrg,
          changed: changedSettings(settingsForOrg(storedRef.current, currentOrg), newSettings),
        });
      }
      const next = withSettingsChange(storedRef.current, newSettings, currentOrg);
      storedRef.current = next;
      setStored(next);
      if (loadedRef.current) persist(next);
      else void loadSettings();
    },
    [loadSettings],
  );

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
