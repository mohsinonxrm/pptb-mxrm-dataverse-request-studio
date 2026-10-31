// Run history (#226): what's kept, how it's capped, and the hook that
// replaces each mode's in-memory Recent runs.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { RecentRun } from './readState';

const settings = vi.hoisted(() => ({ keepRunHistory: true }));
vi.mock('../host/usePersistedSettings', () => ({
  usePersistedSettings: () => [settings, () => {}],
}));

import {
  MAX_HISTORY,
  MAX_SNAPSHOT_CHARS,
  RECENT_SHOWN,
  clearAllHistory,
  historyEntry,
  mergeRuns,
  modeHistory,
  readHistory,
  recordRuns,
  runHistoryKey,
  useRecentRuns,
} from './runHistory';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const run = (n: number, over: Partial<RecentRun> = {}): RecentRun => ({
  id: `r-${n}`,
  modeId: 'retrieve-multiple',
  url: `/api/data/v9.2/accounts?$top=${n}`,
  method: 'GET',
  ts: 1_000 + n,
  status: 200,
  ms: 10,
  rowCount: n,
  snapshot: { table: 'account', top: n, headers: [] },
  ...over,
});

beforeEach(() => {
  localStorage.clear();
  settings.keepRunHistory = true;
});

describe('historyEntry', () => {
  it('drops credential and impersonation headers from the snapshot, at any depth', () => {
    const headers = [
      { name: 'Prefer', value: 'odata.maxpagesize=5', enabled: true },
      { name: 'Authorization', value: 'Bearer secret', enabled: true },
      { name: 'MSCRMCallerID', value: '00000000-0000-0000-0000-000000000009', enabled: true },
      { name: 'callerobjectid', value: 'abc', enabled: false },
    ];
    const kept = historyEntry(
      run(1, { snapshot: { headers, groups: [{ parts: [{ headers }] }] } }),
    );
    const text = JSON.stringify(kept);
    expect(text).not.toMatch(/Authorization|secret|MSCRMCallerID|callerobjectid|0009/);
    expect(text).toContain('odata.maxpagesize=5');
  });

  it('keeps a run whose snapshot is too large, without the snapshot', () => {
    const big = run(1, { snapshot: { rows: 'x'.repeat(MAX_SNAPSHOT_CHARS + 1) } });
    const kept = historyEntry(big);
    expect(kept.snapshot).toBeUndefined();
    expect(kept.id).toBe('r-1');
  });
});

describe('mergeRuns / recordRuns / readHistory', () => {
  it('adds or replaces by id, newest first, capped', () => {
    const merged = mergeRuns([run(1), run(2)], [run(3), { ...run(1), rowCount: 99 }]);
    expect(merged.map((r) => r.id)).toEqual(['r-3', 'r-2', 'r-1']);
    expect(merged[2].rowCount).toBe(99);
    const many = Array.from({ length: MAX_HISTORY + 5 }, (_, i) => run(i));
    expect(mergeRuns([], many)).toHaveLength(MAX_HISTORY);
    expect(mergeRuns([], many)[0].id).toBe(`r-${MAX_HISTORY + 4}`);
  });

  it('persists per environment and reads back', () => {
    recordRuns([run(1)], 'org1.crm.dynamics.com');
    recordRuns([run(2)], 'org2.crm.dynamics.com');
    expect(readHistory('org1.crm.dynamics.com').map((r) => r.id)).toEqual(['r-1']);
    expect(readHistory('org2.crm.dynamics.com').map((r) => r.id)).toEqual(['r-2']);
    clearAllHistory();
    expect(readHistory('org1.crm.dynamics.com')).toEqual([]);
    expect(readHistory('org2.crm.dynamics.com')).toEqual([]);
  });

  it('reads unreadable or foreign data as empty', () => {
    localStorage.setItem(runHistoryKey('x'), '{not json');
    expect(readHistory('x')).toEqual([]);
    localStorage.setItem(runHistoryKey('x'), JSON.stringify([{ id: 1 }, run(4)]));
    expect(readHistory('x').map((r) => r.id)).toEqual(['r-4']);
  });
});

describe('useRecentRuns', () => {
  let root: Root | null = null;
  let host: HTMLDivElement | null = null;
  let api: ReturnType<typeof useRecentRuns> | null = null;
  function Probe({ bucket }: { bucket: string }) {
    api = useRecentRuns(bucket);
    return null;
  }
  const mount = (bucket = 'retrieve-multiple') => {
    host = document.createElement('div');
    root = createRoot(host);
    act(() => root!.render(<Probe bucket={bucket} />));
  };
  afterEach(() => {
    act(() => root?.unmount());
    root = null;
    api = null;
  });
  // What every mode does on Execute.
  const execute = (r: RecentRun) => act(() => api![1]((rs) => [r, ...rs].slice(0, RECENT_SHOWN)));

  it('lists the latest runs and keeps the older ones the mode slices off', () => {
    mount();
    for (let i = 1; i <= RECENT_SHOWN + 2; i++) execute(run(i));
    expect(api![0]).toHaveLength(RECENT_SHOWN);
    expect(api![0][0].id).toBe(`r-${RECENT_SHOWN + 2}`);
    expect(modeHistory('retrieve-multiple')).toHaveLength(RECENT_SHOWN + 2);
  });

  it('survives a remount (a reload), per mode', () => {
    mount();
    execute(run(1));
    act(() => root!.unmount());
    mount();
    expect(api![0].map((r) => r.id)).toEqual(['r-1']);
    act(() => root!.unmount());
    mount('create');
    expect(api![0]).toEqual([]);
  });

  it('updates a run in place (e.g. its row count after Retrieve all)', () => {
    mount();
    execute(run(1));
    act(() => api![1]((rs) => rs.map((r) => (r.id === 'r-1' ? { ...r, rowCount: 500 } : r))));
    expect(api![0][0].rowCount).toBe(500);
    expect(modeHistory('retrieve-multiple')[0].rowCount).toBe(500);
  });

  it('writes nothing with history off, but still lists the session’s runs', () => {
    settings.keepRunHistory = false;
    mount();
    execute(run(1));
    expect(api![0].map((r) => r.id)).toEqual(['r-1']);
    expect(localStorage.length).toBe(0);
  });

  it('empties when the history is cleared', () => {
    mount();
    execute(run(1));
    act(() => clearAllHistory());
    expect(api![0]).toEqual([]);
  });

  it('keeps this session’s full snapshot in memory; storage gets the stripped one', () => {
    mount();
    const headers = [{ name: 'MSCRMCallerID', value: 'u1', enabled: true }];
    execute(run(1, { snapshot: { headers } }));
    expect(JSON.stringify(api![0][0].snapshot)).toContain('MSCRMCallerID');
    expect(JSON.stringify(modeHistory('retrieve-multiple')[0].snapshot)).not.toContain(
      'MSCRMCallerID',
    );
  });
});
