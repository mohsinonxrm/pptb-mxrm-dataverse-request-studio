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
  __resetHistoryCache,
  clearAllHistory,
  fitBudget,
  flushHistory,
  historyEntry,
  mergeRuns,
  modeHistory,
  readHistory,
  recordRuns,
  runHistoryKey,
  setHistoryEnabled,
  useRecentRuns,
} from './runHistory';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const NO_CALLER = '00000000-0000-0000-0000-000000000000';
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
/** What a reload sees: storage, not the in-memory copy. */
const stored = (scope?: string) => {
  flushHistory();
  __resetHistoryCache();
  return readHistory(scope);
};

beforeEach(() => {
  __resetHistoryCache();
  localStorage.clear();
  settings.keepRunHistory = true;
});

describe('historyEntry', () => {
  it('drops credentials and resets caller ids to the off preset, at any depth', () => {
    const headers = [
      { name: 'Prefer', value: 'odata.maxpagesize=5', enabled: true },
      { name: 'Authorization', value: 'Bearer secret', enabled: true },
      { name: 'MSCRMCallerID', value: '00000000-0000-0000-0000-000000000009', enabled: true },
      { name: 'callerobjectid', value: NO_CALLER, enabled: false },
    ];
    const kept = historyEntry(
      run(1, { snapshot: { headers, groups: [{ parts: [{ headers }] }] } }),
    );
    const text = JSON.stringify(kept);
    expect(text).not.toMatch(/Authorization|secret|0009/);
    expect(text).toContain('odata.maxpagesize=5');
    // The rows stay, off and blank, so a restore has its presets back.
    expect((kept.snapshot as { headers: unknown[] }).headers).toEqual([
      headers[0],
      { name: 'MSCRMCallerID', value: NO_CALLER, enabled: false },
      headers[3],
    ]);
    expect(kept.impersonationCleared).toBe(true);
  });

  it('only marks a run whose caller id was actually removed', () => {
    const presets = [{ name: 'MSCRMCallerID', value: NO_CALLER, enabled: false }];
    expect(historyEntry(run(1, { snapshot: { headers: presets } })).impersonationCleared).toBe(
      undefined,
    );
  });

  it('keeps a run whose snapshot is too large, without the snapshot', () => {
    const kept = historyEntry(run(1, { snapshot: { rows: 'x'.repeat(MAX_SNAPSHOT_CHARS + 1) } }));
    expect(kept.snapshot).toBeUndefined();
    expect(kept.id).toBe('r-1');
  });
});

describe('mergeRuns / fitBudget', () => {
  it('adds or replaces by mode and id, newest first, capped', () => {
    const merged = mergeRuns([run(1), run(2)], [run(3), { ...run(1), rowCount: 99 }]);
    expect(merged.map((r) => r.id)).toEqual(['r-3', 'r-2', 'r-1']);
    expect(merged[2].rowCount).toBe(99);
    // Same id from another mode (same millisecond) is another run.
    expect(mergeRuns([run(1)], [run(1, { modeId: 'create' })])).toHaveLength(2);
    const many = Array.from({ length: MAX_HISTORY + 5 }, (_, i) => run(i));
    expect(mergeRuns([], many)).toHaveLength(MAX_HISTORY);
    expect(mergeRuns([], many)[0].id).toBe(`r-${MAX_HISTORY + 4}`);
  });

  it('over budget, drops the oldest snapshots first, then the oldest runs', () => {
    const runs = [3, 2, 1].map((n) => run(n, { snapshot: { blob: 'x'.repeat(1000) } }));
    const one = JSON.stringify(runs[0]).length;
    const trimmed = fitBudget(runs, one + 700);
    expect(trimmed.map((r) => r.id)).toEqual(['r-3', 'r-2', 'r-1']);
    expect(trimmed.map((r) => r.snapshot !== undefined)).toEqual([true, false, false]);
    expect(fitBudget(runs, 300).length).toBeLessThan(3);
  });
});

describe('recordRuns / readHistory', () => {
  it('persists per environment and reads back after a reload', () => {
    recordRuns([run(1)], 'org1.crm.dynamics.com');
    recordRuns([run(2)], 'org2.crm.dynamics.com');
    expect(stored('org1.crm.dynamics.com').map((r) => r.id)).toEqual(['r-1']);
    expect(readHistory('org2.crm.dynamics.com').map((r) => r.id)).toEqual(['r-2']);
    clearAllHistory();
    expect(stored('org1.crm.dynamics.com')).toEqual([]);
    expect(readHistory('org2.crm.dynamics.com')).toEqual([]);
  });

  it('writes nothing while history is off', () => {
    setHistoryEnabled(false);
    recordRuns([run(1)], 'org1');
    expect(stored('org1')).toEqual([]);
  });

  it('reads unreadable or foreign data as empty', () => {
    localStorage.setItem(runHistoryKey('x'), '{not json');
    expect(readHistory('x')).toEqual([]);
    __resetHistoryCache();
    localStorage.setItem(runHistoryKey('x'), JSON.stringify([{ id: 1 }, run(4)]));
    expect(readHistory('x').map((r) => r.id)).toEqual(['r-4']);
  });

  it('a full quota keeps a smaller history rather than none', () => {
    const big = Array.from({ length: 50 }, (_, i) =>
      run(i, { snapshot: { blob: 'x'.repeat(30_000) } }),
    );
    const setItem = Storage.prototype.setItem;
    const spy = vi.spyOn(Storage.prototype, 'setItem').mockImplementation(function (
      this: Storage,
      k: string,
      v: string,
    ) {
      if (v.length > 400_000) throw new DOMException('full', 'QuotaExceededError');
      return setItem.call(this, k, v);
    });
    recordRuns(big, 'q');
    const kept = stored('q');
    spy.mockRestore();
    expect(kept).toHaveLength(50); // every run, some without their snapshot
    expect(kept.filter((r) => r.snapshot === undefined).length).toBeGreaterThan(0);
    expect(JSON.stringify(kept).length).toBeLessThanOrEqual(400_000);
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
  const reload = (bucket?: string) => {
    act(() => root!.unmount());
    flushHistory();
    __resetHistoryCache();
    mount(bucket);
  };
  afterEach(() => {
    act(() => root?.unmount());
    root = null;
    api = null;
  });
  // What every mode does on Execute.
  const execute = (r: RecentRun, setter = api![1]) =>
    act(() => setter((rs) => [r, ...rs].slice(0, RECENT_SHOWN)));

  it('lists the latest runs and keeps the older ones the mode slices off', () => {
    mount();
    for (let i = 1; i <= RECENT_SHOWN + 2; i++) execute(run(i));
    expect(api![0]).toHaveLength(RECENT_SHOWN);
    expect(api![0][0].id).toBe(`r-${RECENT_SHOWN + 2}`);
    expect(modeHistory('retrieve-multiple')).toHaveLength(RECENT_SHOWN + 2);
  });

  it('survives a reload, per mode', () => {
    mount();
    execute(run(1));
    reload();
    expect(api![0].map((r) => r.id)).toEqual(['r-1']);
    reload('create');
    expect(api![0]).toEqual([]);
  });

  it('updates a run in place (e.g. its row count after Retrieve all)', () => {
    mount();
    execute(run(1));
    act(() => api![1]((rs) => rs.map((r) => (r.id === 'r-1' ? { ...r, rowCount: 500 } : r))));
    expect(api![0][0].rowCount).toBe(500);
    expect(stored()[0].rowCount).toBe(500);
  });

  it('with history off: lists the session’s runs and writes nothing', () => {
    settings.keepRunHistory = false;
    mount();
    execute(run(1));
    expect(api![0].map((r) => r.id)).toEqual(['r-1']);
    flushHistory();
    expect(localStorage.length).toBe(0);
  });

  it('a run that finishes after history was turned off is not written', () => {
    mount();
    const setter = api![1]; // the setter an in-flight Execute holds
    act(() => {
      setHistoryEnabled(false);
      clearAllHistory();
    });
    execute(run(1), setter);
    expect(stored()).toEqual([]);
  });

  it('keeps listing new runs when storage can’t be written', () => {
    const spy = vi.spyOn(Storage.prototype, 'setItem').mockImplementation(() => {
      throw new DOMException('full', 'QuotaExceededError');
    });
    mount();
    execute(run(1));
    act(() => flushHistory());
    execute(run(2));
    spy.mockRestore();
    expect(api![0].map((r) => r.id)).toEqual(['r-2', 'r-1']);
  });

  it('empties when the history is cleared', () => {
    mount();
    execute(run(1));
    act(() => clearAllHistory());
    expect(api![0]).toEqual([]);
  });

  it('keeps this session’s full snapshot in memory; storage gets the reset one', () => {
    mount();
    const headers = [{ name: 'MSCRMCallerID', value: 'u1', enabled: true }];
    execute(run(1, { snapshot: { headers } }));
    expect(JSON.stringify(api![0][0].snapshot)).toContain('u1');
    const kept = stored()[0];
    expect(JSON.stringify(kept.snapshot)).not.toContain('u1');
    expect(kept.impersonationCleared).toBe(true);
  });
});
