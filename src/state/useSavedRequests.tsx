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
//   • `save(entry, opts)`  — upsert by id (insert if new, overwrite if id matches);
//                            refused when a new entry would exceed MAX_SAVED,
//                            unless `opts.replaceId` names an entry to drop
//   • `capacity`           — { max, full, oldest } for the library-full prompt
//   • `rename(id, name)`   — set the name field
//   • `remove(id)`         — drop one entry by id
//   • `clearAll()`         — wipe the entire store (used by a "reset" affordance)
//   • `importEntries(list)`— merge entries from an export file (newer wins by id)
//   • `findById(id)`       — read a single entry
//   • `findByName(name)`   — case-insensitive name lookup, for collision detection
//   • `error`              — last write error message, null if clean

import { useCallback, useEffect, useMemo, useState, useSyncExternalStore } from 'react';
import {
  loadSaved,
  persistSaved,
  MAX_SAVED,
  SAVED_REQUESTS_KEY_PREFIX,
  SAVED_SCHEMA_VERSION,
  getOrgScope,
  getSavedRequestsKey,
  hashState,
  type SavedRequest,
} from './savedRequests';
import { stripCredentialRows } from '../engine/sensitiveHeaders';

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

export interface SaveResult {
  ok: boolean;
  error?: string;
  /** True when the save was refused because the library is at MAX_SAVED. */
  full?: boolean;
}

export interface ImportResult extends SaveResult {
  added: number;
  /** Existing entries (same id) replaced by a newer copy from the file. */
  updated: number;
  /** Already present, identical or newer here. */
  unchanged: number;
  /** New entries left out because the library reached MAX_SAVED. */
  skippedFull: number;
}

export interface SavedCapacity {
  /** MAX_SAVED — entries allowed per org scope, across all modes. */
  max: number;
  /** Entries in this org scope's library right now. */
  count: number;
  /** True when a NEW entry can't be added without removing one. */
  full: boolean;
  /** Least recently saved entry (any mode), the default replacement offer. */
  oldest: SavedRequest | undefined;
}

export interface UseSavedRequestsApi {
  saved: SavedRequest[];
  save: (entry: SavedRequest, opts?: { replaceId?: string }) => SaveResult;
  capacity: SavedCapacity;
  rename: (id: string, name: string) => void;
  remove: (id: string) => void;
  clearAll: () => void;
  importEntries: (entries: SavedRequest[]) => ImportResult;
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

  const writeBack = useCallback((next: SavedRequest[]): SaveResult => {
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
  //   • Else → prepend, unless the library is at MAX_SAVED. Then the save is
  //     refused (`full: true`) — the caller asks the user what to remove and
  //     retries with `replaceId`. Entries are never evicted silently (#100):
  //     the library is shared by every mode, so the oldest entry is often
  //     one the user can't even see from the current mode's list.
  const save = useCallback(
    (input: SavedRequest, opts?: { replaceId?: string }): SaveResult => {
      // Credential headers are never stored (#217).
      const entry: SavedRequest = {
        ...input,
        state: stripCredentialRows(input.state),
        schemaVersion: SAVED_SCHEMA_VERSION,
      };
      const idx = list.findIndex((e) => e.id === entry.id);
      let next: SavedRequest[];
      if (idx >= 0) {
        next = [...list];
        next[idx] = entry;
      } else {
        const kept = opts?.replaceId ? list.filter((e) => e.id !== opts.replaceId) : list;
        if (kept.length >= MAX_SAVED) {
          return {
            ok: false,
            full: true,
            error: `The saved-request library is full (${MAX_SAVED} for this connection).`,
          };
        }
        next = [entry, ...kept];
      }
      return writeBack(next);
    },
    [list, writeBack],
  );

  const capacity = useMemo<SavedCapacity>(
    () => ({
      max: MAX_SAVED,
      count: list.length,
      full: list.length >= MAX_SAVED,
      oldest: list.reduce<SavedRequest | undefined>(
        (o, e) => (!o || e.savedAt < o.savedAt ? e : o),
        undefined,
      ),
    }),
    [list],
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

  // Import: entries land in the CURRENT org's library (re-tagged with its
  // scope). Same id → the copy with the later savedAt wins; new ids are
  // added until MAX_SAVED, the rest reported as skipped — never evicting.
  const importEntries = useCallback(
    (entries: SavedRequest[]): ImportResult => {
      const scope = getOrgScope();
      const next = [...list];
      const r = { added: 0, updated: 0, unchanged: 0, skippedFull: 0 };
      for (const raw of entries) {
        const entry: SavedRequest = {
          ...raw,
          state: stripCredentialRows(raw.state),
          orgScope: scope,
          schemaVersion: raw.schemaVersion ?? 1,
        };
        const idx = next.findIndex((e) => e.id === entry.id);
        if (idx >= 0) {
          const cur = next[idx];
          const same = cur.name === entry.name && hashState(cur.state) === hashState(entry.state);
          if (!same && entry.savedAt > cur.savedAt) {
            next[idx] = entry;
            r.updated++;
          } else {
            r.unchanged++;
          }
        } else if (next.length >= MAX_SAVED) {
          r.skippedFull++;
        } else {
          next.push(entry);
          r.added++;
        }
      }
      if (r.added === 0 && r.updated === 0) return { ok: true, ...r };
      next.sort((a, b) => b.savedAt - a.savedAt);
      return { ...writeBack(next), ...r };
    },
    [list, writeBack],
  );

  const findById = useCallback((id: string) => list.find((e) => e.id === id), [list]);
  const findByName = useCallback(
    (name: string) => list.find((e) => e.name.toLowerCase() === name.toLowerCase()),
    [list],
  );

  return {
    saved: list,
    save,
    capacity,
    rename,
    remove,
    clearAll,
    importEntries,
    findById,
    findByName,
    error,
  };
}
