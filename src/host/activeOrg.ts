// Active org — which Dataverse environment the in-memory caches belong to.
//
// PPTB can switch a running tool to another connection. The host then sends
// `connection:updated`, and the bridge re-emits the session with the new
// `environmentUrl`. Metadata caches (cache.ts, the live table registry,
// cascade config) are keyed by this org, so after a switch every read
// misses and refetches against the new environment, while switching back
// finds the old org's data still warm (#52).
//
// The key is the environment host, lower-cased (`contoso.crm.dynamics.com`).
// Empty when there is no connection (standalone, before the handshake, or
// after the connection is deleted).

import { useSyncExternalStore } from 'react';
import { subscribe, type HostSession } from './pptbBridge';

type OrgListener = (org: string) => void;

const listeners = new Set<OrgListener>();
let current = '';

function orgKeyOf(s: HostSession): string {
  if (!s.embedded || !s.connected || !s.environmentUrl) return '';
  return s.environmentUrl
    .trim()
    .replace(/^https?:\/\//i, '')
    .replace(/\/+$/, '')
    .toLowerCase();
}

// Runs synchronously inside the bridge's emit, before React re-renders, so
// any read made while rendering for the new session already sees the new org.
subscribe((s) => {
  const next = orgKeyOf(s);
  if (next === current) return;
  current = next;
  for (const l of listeners) {
    try {
      l(next);
    } catch (e) {
      console.error('[activeOrg] listener threw', e);
    }
  }
});

/** The active org key. Empty when there is no connection. */
export const activeOrg = (): string => current;

/** Fires after the active org changes. Returns an unsubscribe function. */
export function onActiveOrgChange(cb: OrgListener): () => void {
  listeners.add(cb);
  return () => {
    listeners.delete(cb);
  };
}

/** React binding — re-renders when the active org changes. */
export function useActiveOrg(): string {
  return useSyncExternalStore(onActiveOrgChange, activeOrg);
}
