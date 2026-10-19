// In-memory drafts of each mode's request, so switching modes in the
// sidebar doesn't throw away an unsaved request (#101).
//
// ModeRouter unmounts the active mode on every switch and each mode keeps
// its request in local `useState`. Instead of keeping modes mounted (their
// effects and fetches would keep running), a mode stashes its serialized
// request — the same snapshot Save writes — when it unmounts, and starts
// from it on the next mount, via the same deserializer as loading a saved
// request. The Save baseline (lastSavedId / lastSavedHash) travels with it,
// so a restored draft is exactly as dirty as when it was left.
//
// Scope: this session only (module memory, never persisted), keyed by mode
// key AND org. A draft is only kept for the org it was made against: on a
// connection switch every draft is dropped, so nothing prepared against one
// environment comes back against another (#93).

import { useEffect, useRef, useState } from 'react';
import { getOrgScope } from './savedRequests';

export interface ModeDraft<T = unknown> {
  state: T;
  lastSavedId: string | undefined;
  lastSavedHash: string | null;
}

let draftsOrg: string | null = null;
const drafts = new Map<string, ModeDraft>();

/** Drop every draft when the active org isn't the one they were made for. */
function syncOrg(org: string): void {
  if (draftsOrg === org) return;
  drafts.clear();
  draftsOrg = org;
}

export function readModeDraft(key: string): ModeDraft | undefined {
  syncOrg(getOrgScope());
  return drafts.get(key);
}

/** Store a draft made against `org`; ignored when that is no longer the active org. */
export function writeModeDraft(org: string, key: string, draft: ModeDraft): void {
  const current = getOrgScope();
  syncOrg(current);
  if (org === current) drafts.set(key, draft);
}

export function clearModeDrafts(): void {
  drafts.clear();
}

/**
 * Mount-time initial values for a mode: the stashed draft for `key` (run
 * through `deserialize`), or `init()` with no Save baseline. Read once —
 * pass the result straight to the mode's `useState` calls.
 */
export function useModeDraft<S, T>(
  key: string,
  deserialize: (snap: S) => T,
  init: () => T,
): ModeDraft<T> {
  const [initial] = useState<ModeDraft<T>>(() => {
    const d = readModeDraft(key);
    if (d) {
      try {
        return {
          state: deserialize(d.state as S),
          lastSavedId: d.lastSavedId,
          lastSavedHash: d.lastSavedHash,
        };
      } catch {
        // A draft that can't be rebuilt is dropped; start fresh.
      }
    }
    return { state: init(), lastSavedId: undefined, lastSavedHash: null };
  });
  return initial;
}

/**
 * Stash the mode's current request (serialized) when it unmounts or its
 * `key` changes. The draft is tagged with the first org this instance saw
 * connected, so a re-render after a connection switch can't move it to the
 * new org.
 */
export function useStashModeDraft(
  key: string,
  serialized: unknown,
  lastSavedId: string | undefined,
  lastSavedHash: string | null,
): void {
  const org = useRef(connectedOrg());
  const latest = useRef<ModeDraft>({ state: serialized, lastSavedId, lastSavedHash });
  useEffect(() => {
    if (!org.current) org.current = connectedOrg();
    latest.current = { state: serialized, lastSavedId, lastSavedHash };
  });
  useEffect(() => () => writeModeDraft(org.current || getOrgScope(), key, latest.current), [key]);
}

function connectedOrg(): string {
  const org = getOrgScope();
  return org === 'default' ? '' : org;
}
