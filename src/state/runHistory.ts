// Run history (#226): each mode's Recent runs, kept across reloads and
// ToolBox restarts.
//
// One localStorage bucket per environment, like the saved library (#105:
// a tool's localStorage survives updates and restarts), holding the last
// MAX_HISTORY runs of every mode. An entry is a RecentRun: the request
// snapshot (the same one Saved stores), status, row count, duration and
// time. Never the response body, and never credential or impersonation
// headers (stripHistoryRows). A snapshot too large to keep is dropped; the
// run is still listed, it just can't be restored.
//
// `useRecentRuns(bucket)` is a drop-in for a mode's
// `useState<RecentRun[]>([])`: it returns the mode's latest RECENT_SHOWN runs
// and a setter that takes the same updaters the modes already pass. Runs the
// updater drops (a mode's `.slice(0, 8)`) stay in the history; only Clear
// removes them. With `keepRunHistory` off nothing is written, and Recent
// runs last for the session, as before.

import { useCallback, useEffect, useRef, useState, type SetStateAction } from 'react';
import type { RecentRun } from './readState';
import { getOrgScope } from './savedRequests';
import { stripHistoryRows } from '../engine/sensitiveHeaders';
import { usePersistedSettings } from '../host/usePersistedSettings';
import { useHostSession } from '../host/HostContext';

export const RUN_HISTORY_KEY_PREFIX = 'drs:run-history:v1';
/** Runs kept per environment, across all modes. */
export const MAX_HISTORY = 200;
/** Runs a mode lists in its sidebar. */
export const RECENT_SHOWN = 8;
/** A snapshot larger than this (as JSON) isn't kept: a big Bulk Write
 *  import would otherwise crowd out the rest of the history. */
export const MAX_SNAPSHOT_CHARS = 64 * 1024;

export const runHistoryKey = (scope = getOrgScope()) => `${RUN_HISTORY_KEY_PREFIX}:${scope}`;

function isRun(x: unknown): x is RecentRun {
  const r = x as RecentRun;
  return (
    !!r &&
    typeof r === 'object' &&
    typeof r.id === 'string' &&
    typeof r.modeId === 'string' &&
    typeof r.ts === 'number' &&
    typeof r.url === 'string' &&
    typeof r.method === 'string' &&
    typeof r.status === 'number' &&
    typeof r.ms === 'number'
  );
}

/** The environment's history, newest first. Unreadable data reads as empty. */
export function readHistory(scope = getOrgScope()): RecentRun[] {
  try {
    const raw = localStorage.getItem(runHistoryKey(scope));
    if (!raw) return [];
    const list: unknown = JSON.parse(raw);
    return Array.isArray(list) ? list.filter(isRun).sort((a, b) => b.ts - a.ts) : [];
  } catch {
    return [];
  }
}

/** What a run keeps: no sensitive headers, and no snapshot when it's too large. */
export function historyEntry(run: RecentRun): RecentRun {
  if (run.snapshot === undefined) return run;
  const snapshot = stripHistoryRows(run.snapshot);
  let size = Infinity;
  try {
    size = JSON.stringify(snapshot)?.length ?? 0;
  } catch {
    // Unserializable: treat as too large.
  }
  const { snapshot: _drop, ...rest } = run;
  return size <= MAX_SNAPSHOT_CHARS ? { ...rest, snapshot } : rest;
}

/** `runs` added to (or replacing, by id) `history`; newest first, capped. */
export function mergeRuns(history: readonly RecentRun[], runs: readonly RecentRun[]): RecentRun[] {
  const byId = new Map(history.map((r) => [r.id, r]));
  for (const r of runs) byId.set(r.id, r);
  return [...byId.values()].sort((a, b) => b.ts - a.ts).slice(0, MAX_HISTORY);
}

const listeners = new Set<() => void>();
const notify = () => listeners.forEach((fn) => fn());

/** Called when any history changes (a run recorded, a clear). */
export function subscribeHistory(fn: () => void): () => void {
  listeners.add(fn);
  return () => listeners.delete(fn);
}

/**
 * Add or update runs in the environment's history. A full storage quota
 * drops the oldest half and tries once more; after that the runs stay in
 * the session only.
 */
export function recordRuns(runs: readonly RecentRun[], scope = getOrgScope()): void {
  let next = mergeRuns(readHistory(scope), runs.map(historyEntry));
  for (let attempt = 0; attempt < 2; attempt++) {
    try {
      localStorage.setItem(runHistoryKey(scope), JSON.stringify(next));
      break;
    } catch {
      next = next.slice(0, Math.ceil(next.length / 2));
    }
  }
  notify();
}

/** Remove every environment's history. */
export function clearAllHistory(): void {
  try {
    const keys: string[] = [];
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i);
      if (k?.startsWith(`${RUN_HISTORY_KEY_PREFIX}:`)) keys.push(k);
    }
    keys.forEach((k) => localStorage.removeItem(k));
  } catch {
    // Storage unavailable: nothing was kept either.
  }
  notify();
}

/** A mode's runs in the environment's history, newest first. */
export const modeHistory = (bucket: string, scope = getOrgScope()) =>
  readHistory(scope).filter((r) => r.modeId === bucket);

/**
 * A mode's Recent runs (#226), in place of `useState<RecentRun[]>([])`.
 * `bucket` is the `modeId` the mode stamps on its runs.
 */
export function useRecentRuns(
  bucket: string,
): [RecentRun[], (action: SetStateAction<RecentRun[]>) => void] {
  const [{ keepRunHistory }] = usePersistedSettings();
  // Re-read when the connection changes: each environment has its own history.
  const { environmentUrl } = useHostSession();
  const scope = getOrgScope();
  const [runs, setRuns] = useState<RecentRun[]>(() =>
    keepRunHistory ? modeHistory(bucket, scope) : [],
  );
  const runsRef = useRef(runs);
  runsRef.current = runs;
  const shownScope = useRef(scope);

  useEffect(() => {
    const scopeChanged = shownScope.current !== scope;
    shownScope.current = scope;
    if (!keepRunHistory) {
      // Session-only runs belong to the environment they ran against.
      if (scopeChanged) setRuns([]);
      return;
    }
    // A run made this session keeps its full snapshot in memory; the stored
    // copy lacks impersonation headers. Anything cleared from storage goes.
    const reload = () => {
      const mine = new Map(runsRef.current.map((r) => [r.id, r]));
      setRuns(modeHistory(bucket, scope).map((r) => mine.get(r.id) ?? r));
    };
    reload();
    return subscribeHistory(reload);
  }, [bucket, scope, environmentUrl, keepRunHistory]);

  const update = useCallback(
    (action: SetStateAction<RecentRun[]>) => {
      // The mode's updater sees what it lists; what it drops stays in history.
      const shown = runsRef.current.slice(0, RECENT_SHOWN);
      const next = typeof action === 'function' ? action(shown) : action;
      const merged = mergeRuns(runsRef.current, next);
      runsRef.current = merged;
      setRuns(merged);
      if (keepRunHistory) recordRuns(next, scope);
    },
    [keepRunHistory, scope],
  );

  return [runs.slice(0, RECENT_SHOWN), update];
}

/** Every run of a mode in this environment, for the History dialog. */
export function useModeHistory(bucket: string, sessionRuns: readonly RecentRun[]): RecentRun[] {
  const [{ keepRunHistory }] = usePersistedSettings();
  const scope = getOrgScope();
  const [runs, setRuns] = useState<RecentRun[]>(() => modeHistory(bucket, scope));
  useEffect(() => {
    const reload = () => setRuns(modeHistory(bucket, scope));
    reload();
    return subscribeHistory(reload);
  }, [bucket, scope]);
  // History off: only this session's runs exist.
  return keepRunHistory ? runs : [...sessionRuns];
}
