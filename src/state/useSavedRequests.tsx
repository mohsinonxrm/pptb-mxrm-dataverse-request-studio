// React hook over the saved-request library (localStorage-backed).
//
// **Module-level singleton store**: every `useSavedRequests()` consumer
// subscribes to the SAME snapshot via `useSyncExternalStore`. When one
// component saves/renames/removes, all other consumers immediately see
// the update — no manual refresh, no mode switch dance.
//
// The previous hook-local `useState` design isolated each consumer's
// list, so the SaveButton (in the FrameHeader) and the SavedLibraryButton
// (also in the FrameHeader) each held their own copy. Saving in one
// didn't show up in the other until something forced a re-mount.
//
// localStorage is read once at module init and again on cross-tab
// `storage` events. Writes go through `persistSaved` then notify all
// subscribers via the same broadcast path.
//
// Returned API:
//   • `saved`              — current list, newest first
//   • `save(entry)`        — upsert by id (insert if new, overwrite if id matches)
//   • `rename(id, name)`   — set the name field
//   • `remove(id)`         — drop one entry by id
//   • `clearAll()`         — wipe the entire store (used by a "reset" affordance)
//   • `findById(id)`       — read a single entry
//   • `findByName(name)`   — case-insensitive name lookup, for collision detection
//   • `error`              — last write error message, null if clean

import { useCallback, useEffect, useState, useSyncExternalStore } from 'react';
import {
  loadSaved,
  persistSaved,
  MAX_SAVED,
  SAVED_REQUESTS_KEY_PREFIX,
  getSavedRequestsKey,
  type SavedRequest,
} from './savedRequests';

// ── Module-level singleton store ────────────────────────────────────

let snapshot: SavedRequest[] = typeof window !== 'undefined' ? loadSaved() : [];
const listeners = new Set<() => void>();

function notify(): void {
  for (const l of listeners) l();
}

function setSnapshot(next: SavedRequest[]): void {
  snapshot = next;
  notify();
}

function subscribe(cb: () => void): () => void {
  listeners.add(cb);
  return () => {
    listeners.delete(cb);
  };
}

function getSnapshot(): SavedRequest[] {
  return snapshot;
}

// Cross-tab sync — when another tab writes ANY saved-requests bucket
// (any org scope), our `storage` listener reloads the snapshot for the
// active org and notifies every local consumer. We match on the prefix
// so writes against the current scope trigger a refresh, and writes
// against other scopes are correctly ignored after the reload.
// Same-tab writes go directly through `setSnapshot` (no `storage` event
// fires for same-tab localStorage writes per the spec).
if (typeof window !== 'undefined') {
  window.addEventListener('storage', (e) => {
    if (e.key && e.key.startsWith(SAVED_REQUESTS_KEY_PREFIX) && e.key === getSavedRequestsKey()) {
      snapshot = loadSaved();
      notify();
    }
  });
}

/**
 * Force-reload the singleton from localStorage. Called when the active
 * Dataverse connection switches — the key we read from changes, so we
 * need to pull from a different bucket. Public so the host bridge or a
 * mode-router effect can trigger it.
 */
export function reloadSavedFromStorage(): void {
  if (typeof window === 'undefined') return;
  const next = loadSaved();
  setSnapshot(next);
}

export interface UseSavedRequestsApi {
  saved: SavedRequest[];
  save: (entry: SavedRequest) => { ok: boolean; error?: string };
  rename: (id: string, name: string) => void;
  remove: (id: string) => void;
  clearAll: () => void;
  findById: (id: string) => SavedRequest | undefined;
  findByName: (name: string) => SavedRequest | undefined;
  /** Last localStorage write error, if any. Cleared on next successful write. */
  error: string | null;
}

export function useSavedRequests(): UseSavedRequestsApi {
  // Subscribe to the singleton store. Every consumer gets the same
  // snapshot; saving in one component immediately renders the new entry
  // in any other component that's also using this hook.
  const list = useSyncExternalStore(subscribe, getSnapshot, getSnapshot);
  // `error` is per-consumer — only the component that triggered a failed
  // write needs to surface it. Successful writes clear it (next call).
  const [error, setError] = useState<string | null>(null);

  // When the active org scope changes, the singleton is pointing at the
  // wrong bucket. Detect the mismatch on every render (cheap string
  // compare) and reload. This covers the case where PPTB pushes a new
  // environmentUrl mid-session without a full page reload.
  const activeKey = typeof window !== 'undefined' ? getSavedRequestsKey() : '';
  useEffect(() => {
    if (!activeKey) return;
    reloadSavedFromStorage();
  }, [activeKey]);

  const writeBack = useCallback((next: SavedRequest[]): { ok: boolean; error?: string } => {
    try {
      persistSaved(next);
      setSnapshot(next); // broadcasts to every subscribed consumer
      setError(null);
      return { ok: true };
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      setError(msg);
      return { ok: false, error: msg };
    }
  }, []);

  // Save (upsert):
  //   • If id already exists in list → overwrite that entry's content + savedAt.
  //   • Else → prepend. If we exceed MAX_SAVED, drop the OLDEST entry
  //     (smallest savedAt). User can manually remove others first if they
  //     want a different eviction order.
  const save = useCallback(
    (entry: SavedRequest) => {
      const idx = list.findIndex((e) => e.id === entry.id);
      let next: SavedRequest[];
      if (idx >= 0) {
        next = [...list];
        next[idx] = entry;
      } else {
        next = [entry, ...list];
        if (next.length > MAX_SAVED) {
          // Sort copy by savedAt asc, drop the oldest until at cap.
          next.sort((a, b) => b.savedAt - a.savedAt);
          next = next.slice(0, MAX_SAVED);
        }
      }
      return writeBack(next);
    },
    [list, writeBack],
  );

  const rename = useCallback(
    (id: string, name: string) => {
      const next = list.map((e) => (e.id === id ? { ...e, name } : e));
      writeBack(next);
    },
    [list, writeBack],
  );

  const remove = useCallback(
    (id: string) => {
      writeBack(list.filter((e) => e.id !== id));
    },
    [list, writeBack],
  );

  const clearAll = useCallback(() => {
    writeBack([]);
  }, [writeBack]);

  const findById = useCallback((id: string) => list.find((e) => e.id === id), [list]);
  const findByName = useCallback(
    (name: string) => list.find((e) => e.name.toLowerCase() === name.toLowerCase()),
    [list],
  );

  return { saved: list, save, rename, remove, clearAll, findById, findByName, error };
}
