// React context for the global Save UI.
//
// The Save button + Saved-library button live in `FrameHeader` (top-right
// toolbar, per v2.2 design), which is App-level chrome rendered ABOVE the
// active mode. The mode owns the actual state being saved — so we need a
// bridge from the mode's local state to the header's render tree.
//
// Each mode wraps its render with `<SaveContextPublisher value={...}>` to
// publish its current save context. FrameHeader reads the active value
// via `useSaveContext()` and renders the buttons (or nothing if no
// publisher is mounted — e.g. specialty modes that don't have meaningful
// persistable state).

import { createContext, useContext, useEffect, type ReactNode } from 'react';
import type { SavedRequest, SavedModeId } from './savedRequests';

export interface SaveContextValue {
  /**
   * Current state snapshot (already serialized — no `dirty` Set).
   * Type-erased so this context stays mode-agnostic; each mode casts
   * inside its onLoadSaved / serializer.
   */
  state: unknown;
  /** Mode id for routing the saved entry. */
  modeId: SavedModeId;
  /** True when state diverges from the last save (or never saved on this entity). */
  dirty: boolean;
  /** Id of the last saved entry that matches this state, if any. Enables silent overwrite. */
  lastSavedId?: string;
  /** Called after a successful save — the mode updates its dirty baseline. */
  onSaved: (saved: SavedRequest) => void;
  /** Called when the user picks a saved entry to load. The mode validates + hydrates state. */
  onLoadSaved: (entry: SavedRequest) => void;
}

// Two pieces:
//   • `SaveContext` carries the current published value
//   • `SaveSetterContext` carries the setter the mode uses to publish
// Split contexts keep consumers from re-rendering when only the setter changes.

type Setter = (value: SaveContextValue | null) => void;

const SaveContext = createContext<SaveContextValue | null>(null);
const SaveSetterContext = createContext<Setter | null>(null);

/**
 * Provider — mount once at the App level (above FrameHeader and the
 * active mode). Holds a single slot for the currently-active save context.
 */
export function SaveContextRoot({
  value, setValue, children,
}: {
  value: SaveContextValue | null;
  setValue: Setter;
  children: ReactNode;
}) {
  return (
    <SaveSetterContext.Provider value={setValue}>
      <SaveContext.Provider value={value}>
        {children}
      </SaveContext.Provider>
    </SaveSetterContext.Provider>
  );
}

/**
 * Hook for the mode — publishes its save context on every render and
 * clears it on unmount so a stale mode's state doesn't leak to the
 * header after mode switch.
 *
 * Pass `null` from a mode that doesn't support saving (e.g. when the
 * mode is mid-load and has nothing meaningful to save yet) — the
 * FrameHeader buttons will hide.
 */
export function usePublishSaveContext(value: SaveContextValue | null): void {
  const setter = useContext(SaveSetterContext);
  useEffect(() => {
    if (!setter) return;
    setter(value);
    return () => setter(null);
  }, [setter, value]);
}

/**
 * Hook for the consumer (FrameHeader). Returns the currently-published
 * save context, or null if no mode is publishing one right now.
 */
export function useSaveContext(): SaveContextValue | null {
  return useContext(SaveContext);
}
