// Run history (#226): each mode's Recent runs, kept across reloads and
// ToolBox restarts.
//
// One localStorage bucket per environment, like the saved library (#105:
// a tool's localStorage survives updates and restarts), holding the last
// MAX_HISTORY runs of every mode. An entry is a RecentRun: the request
// snapshot (the same one Saved stores), status, row count, duration and
// time. Never the response body. Credential headers are removed and
// impersonation caller ids reset (stripHistoryRows); the run is marked so a
// restore can say so.
//
// Storage is shared with the saved library, so the history keeps to a
// budget (MAX_HISTORY_CHARS): over it, the oldest runs lose their snapshots
// first (they stay listed, without restore), then the oldest runs go.
//
// The parsed history is kept in memory per environment and written to
// storage shortly after it changes, so a Retrieve all that updates its run
// on every page doesn't parse and rewrite storage each time.
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
/** The history's share of storage (characters of JSON), well under the
 *  browser's ~5 M quota so the saved library always has room. */
export const MAX_HISTORY_CHARS = 1_000_000;
/** One snapshot larger than this isn't kept at all. */
export const MAX_SNAPSHOT_CHARS = 64 * 1024;
/** How long changes wait before they're written to storage. */
const WRITE_DELAY_MS = 300;

export const runHistoryKey = (scope = getOrgScope()) => `${RUN_HISTORY_KEY_PREFIX}:${scope}`;

// ── On / off ────────────────────────────────────────────────────────────
// Read at write time, so a run that finishes after history was turned off
// is never written (its mode's setter may be from an earlier render).
let historyEnabled = true;
export function setHistoryEnabled(on: boolean): void {
  historyEnabled = on;
}
export const isHistoryEnabled = () => historyEnabled;

// ── Entries ─────────────────────────────────────────────────────────────

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

const jsonSize = (v: unknown): number => {
  try {
    return JSON.stringify(v)?.length ?? 0;
  } catch {
    return Infinity;
  }
};

/** What a run keeps: no credentials, no caller id, no snapshot when it's too large. */
export function historyEntry(run: RecentRun): RecentRun {
  if (run.snapshot === undefined) return run;
  const { state: snapshot, cleared } = stripHistoryRows(run.snapshot);
  const { snapshot: _drop, ...rest } = run;
  const marked = cleared ? { ...rest, impersonationCleared: true } : rest;
  return jsonSize(snapshot) <= MAX_SNAPSHOT_CHARS ? { ...marked, snapshot } : marked;
}

const runKey = (r: RecentRun) => `${r.modeId}\u0000${r.id}`;

/** `runs` added to (or replacing, by mode + id) `history`; newest first, capped. */
export function mergeRuns(history: readonly RecentRun[], runs: readonly RecentRun[]): RecentRun[] {
  const byKey = new Map(history.map((r) => [runKey(r), r]));
  for (const r of runs) byKey.set(runKey(r), r);
  return [...byKey.values()].sort((a, b) => b.ts - a.ts).slice(0, MAX_HISTORY);
}

/**
 * `runs` (newest first) within `budget` characters: the oldest runs lose
 * their snapshots first, then the oldest runs are dropped.
 */
export function fitBudget(runs: readonly RecentRun[], budget = MAX_HISTORY_CHARS): RecentRun[] {
  const out = [...runs];
  let total = jsonSize(out);
  for (let i = out.length - 1; i >= 0 && total > budget; i--) {
    if (out[i].snapshot === undefined) continue;
    const { snapshot: _drop, ...rest } = out[i];
    total -= jsonSize(out[i]) - jsonSize(rest);
    out[i] = rest;
  }
  while (out.length > 0 && total > budget) {
    total -= jsonSize(out[out.length - 1]) + 1;
    out.pop();
  }
  return out;
}

// ── Store ───────────────────────────────────────────────────────────────

const cache = new Map<string, RecentRun[]>();
const pending = new Set<string>();
let writeTimer: ReturnType<typeof setTimeout> | null = null;

function readStorage(scope: string): RecentRun[] {
  try {
    const raw = localStorage.getItem(runHistoryKey(scope));
    if (!raw) return [];
    const list: unknown = JSON.parse(raw);
    return Array.isArray(list) ? list.filter(isRun).sort((a, b) => b.ts - a.ts) : [];
  } catch {
    return [];
  }
}

/** The environment's history, newest first. Unreadable data reads as empty. */
export function readHistory(scope = getOrgScope()): RecentRun[] {
  let runs = cache.get(scope);
  if (!runs) {
    runs = readStorage(scope);
    cache.set(scope, runs);
  }
  return runs;
}

/** Write pending environments now. A full quota halves the budget and tries again. */
export function flushHistory(): void {
  if (writeTimer) clearTimeout(writeTimer);
  writeTimer = null;
  for (const scope of pending) {
    let budget = MAX_HISTORY_CHARS;
    for (let attempt = 0; attempt < 4; attempt++) {
      const runs = fitBudget(cache.get(scope) ?? [], budget);
      try {
        localStorage.setItem(runHistoryKey(scope), JSON.stringify(runs));
        cache.set(scope, runs);
        break;
      } catch {
        budget = Math.floor(budget / 2);
      }
    }
  }
  pending.clear();
}

if (typeof window !== 'undefined') window.addEventListener('pagehide', flushHistory);

type Change = 'update' | 'clear';
const listeners = new Set<(change: Change) => void>();
const notify = (change: Change) => listeners.forEach((fn) => fn(change));

/** Called when any history changes: a run recorded, or a clear. */
export function subscribeHistory(fn: (change: Change) => void): () => void {
  listeners.add(fn);
  return () => listeners.delete(fn);
}

/** Add or update runs in the environment's history (unless history is off). */
export function recordRuns(runs: readonly RecentRun[], scope = getOrgScope()): void {
  if (!historyEnabled) return;
  cache.set(scope, mergeRuns(readHistory(scope), runs.map(historyEntry)));
  pending.add(scope);
  if (!writeTimer) writeTimer = setTimeout(flushHistory, WRITE_DELAY_MS);
  notify('update');
}

/** Remove every environment's history. */
export function clearAllHistory(): void {
  if (writeTimer) clearTimeout(writeTimer);
  writeTimer = null;
  pending.clear();
  cache.clear();
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
  notify('clear');
}

/** A mode's runs in the environment's history, newest first. */
export const modeHistory = (bucket: string, scope = getOrgScope()) =>
  readHistory(scope).filter((r) => r.modeId === bucket);

// ── Hooks ───────────────────────────────────────────────────────────────

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
    setHistoryEnabled(keepRunHistory);
  }, [keepRunHistory]);

  useEffect(() => {
    const scopeChanged = shownScope.current !== scope;
    shownScope.current = scope;
    if (scopeChanged) runsRef.current = [];
    // Kept runs, with this session's in-memory copies (full snapshots) in
    // place of theirs, plus session runs storage doesn't have. A clear
    // drops them all.
    const reload = (change: Change = 'update') => {
      const stored = keepRunHistory ? modeHistory(bucket, scope) : [];
      if (change === 'clear') {
        runsRef.current = stored;
        setRuns(stored);
        return;
      }
      const mine = new Map(runsRef.current.map((r) => [runKey(r), r]));
      const merged = mergeRuns(
        stored.map((r) => mine.get(runKey(r)) ?? r),
        runsRef.current,
      );
      runsRef.current = merged;
      setRuns(merged);
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
      recordRuns(next, scope);
    },
    [scope],
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
  // This session's runs (full snapshots, and any storage didn't keep) win.
  return keepRunHistory ? mergeRuns(runs, sessionRuns) : [...sessionRuns];
}

/** Tests: forget the in-memory copy, so the next read comes from storage. */
export function __resetHistoryCache(): void {
  if (writeTimer) clearTimeout(writeTimer);
  writeTimer = null;
  pending.clear();
  cache.clear();
  historyEnabled = true;
}
