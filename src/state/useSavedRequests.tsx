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
// Nothing read from storage is lost by a write (#220): items this build
// can't use are written back unchanged after the entries, and a library
// that can't be read at all is never written over — every write is refused
// until the user keeps a copy (`unreadableText`) or starts a new library.
//
// Returned API:
//   • `saved`              — current list, newest first
//   • `save(entry, opts)`  — upsert by id (insert if new, overwrite if id matches);
//                            refused when a new entry would exceed MAX_SAVED,
//                            unless `opts.replaceId` names an entry to drop
//   • `capacity`           — { max, full, oldest } for the library-full prompt
//   • `rename(id, name)`   — set the name field
//   • `remove(id)`         — drop one entry by id
//   • `clearAll()`         — wipe the entries (items this build can't read stay)
//   • `importEntries(list)`— merge entries from an export file (newer wins by id)
//   • `findById(id)`       — read a single entry
//   • `findByName(name)`   — case-insensitive name lookup, for collision detection
//   • `health`             — unreadable library / hidden items, for the library banner
//   • `startNewLibrary()`  — replace an unreadable library with an empty one
//   • `error`              — last write error message, null if clean

import { useCallback, useEffect, useMemo, useState, useSyncExternalStore } from 'react';
import {
  readLibrary,
  backupUnreadable,
  persistSaved,
  MAX_SAVED,
  SAVED_REQUESTS_KEY_PREFIX,
  SAVED_SCHEMA_VERSION,
  getOrgScope,
  getSavedRequestsKey,
  hashState,
  type SavedLibrary,
  scrubStoredCredentials,
  type SavedRequest,
} from './savedRequests';
import { stripCredentialRows } from '../engine/sensitiveHeaders';

// ── Module-level singleton store ────────────────────────────────────

// Credentials saved by an earlier build leave storage once, at startup (#217).
if (typeof window !== 'undefined') scrubStoredCredentials();

interface StoreState extends SavedLibrary {
  /** Where the unreadable text was copied, or null (none, or storage refused). */
  backupKey: string | null;
}

const EMPTY: StoreState = { entries: [], preserved: [], unreadable: null, backupKey: null };

/** Read the current connection's library; back up an unreadable one. */
function readStore(): StoreState {
  const lib = readLibrary();
  const backupKey =
    lib.unreadable !== null ? backupUnreadable(getOrgScope(), lib.unreadable) : null;
  return { ...lib, backupKey };
}

let store: StoreState = typeof window !== 'undefined' ? readStore() : EMPTY;
const listeners = new Set<() => void>();

function notify(): void {
  for (const l of listeners) l();
}

function setStore(next: StoreState): void {
  store = next;
  notify();
}

function subscribe(cb: () => void): () => void {
  listeners.add(cb);
  return () => {
    listeners.delete(cb);
  };
}

function getSnapshot(): StoreState {
  return store;
}

const UNREADABLE_ERROR =
  "The saved requests for this connection couldn't be read, so nothing was saved over them. " +
  'Open Saved requests to keep a copy, or start a new library.';

// Cross-tab sync — when another tab writes ANY saved-requests bucket
// (any org scope), our `storage` listener reloads the snapshot for the
// active org and notifies every local consumer. We match on the prefix
// so writes against the current scope trigger a refresh, and writes
// against other scopes are correctly ignored after the reload.
// Same-tab writes go directly through `setStore` (no `storage` event
// fires for same-tab localStorage writes per the spec).
if (typeof window !== 'undefined') {
  window.addEventListener('storage', (e) => {
    if (e.key && e.key.startsWith(SAVED_REQUESTS_KEY_PREFIX) && e.key === getSavedRequestsKey()) {
      store = readStore();
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
  setStore(readStore());
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

/** What the library banner needs to know about the stored data (#220). */
export interface SavedLibraryHealth {
  /** The stored library isn't readable; every write is refused. */
  unreadable: boolean;
  /** The unreadable stored text, for "Save a copy". */
  unreadableText: string | null;
  /** Whether a backup copy of the unreadable text is in storage. */
  backedUp: boolean;
  /** Stored items this build can't show (damaged, newer version); kept as they are. */
  hidden: number;
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
  health: SavedLibraryHealth;
  startNewLibrary: () => SaveResult;
  /** Last localStorage write error, if any. Cleared on next successful write. */
  error: string | null;
}

export function useSavedRequests(): UseSavedRequestsApi {
  // Subscribe to the singleton store. Every consumer gets the same
  // snapshot; saving in one component immediately renders the new entry
  // in any other component that's also using this hook.
  const snap = useSyncExternalStore(subscribe, getSnapshot, getSnapshot);
  const list = snap.entries;
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
    // Never write over a library that couldn't be read (#220).
    if (store.unreadable !== null) {
      setError(UNREADABLE_ERROR);
      return { ok: false, error: UNREADABLE_ERROR };
    }
    try {
      persistSaved(next, store.preserved);
      setStore({ ...store, entries: next }); // broadcasts to every subscribed consumer
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

  // Replace an unreadable library with an empty one. The unreadable text
  // stays in its backup key (when storage allowed one).
  const startNewLibrary = useCallback((): SaveResult => {
    try {
      persistSaved([]);
      setStore({ ...EMPTY });
      setError(null);
      return { ok: true };
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      setError(msg);
      return { ok: false, error: msg };
    }
  }, []);

  const health = useMemo<SavedLibraryHealth>(
    () => ({
      unreadable: snap.unreadable !== null,
      unreadableText: snap.unreadable,
      backedUp: snap.backupKey !== null,
      hidden: snap.preserved.length,
    }),
    [snap],
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
    health,
    startNewLibrary,
    // The refusal no longer applies once a new library was started from
    // another consumer (the library popover) — don't keep showing it.
    error: error === UNREADABLE_ERROR && snap.unreadable === null ? null : error,
  };
}
