// One Recent-runs push for every mode (#244).
import { afterEach, describe, expect, it, vi } from 'vitest';
import { addRun, RECENT_SHOWN, type NewRun } from './runHistory';
import type { RecentRun } from './readState';

const run = (extra: Partial<NewRun> = {}): NewRun => ({
  modeId: 'create',
  url: '/api/data/v9.2/accounts',
  method: 'POST',
  status: 201,
  ms: 5,
  ...extra,
});

afterEach(() => {
  vi.useRealTimers();
});

describe('addRun', () => {
  it('puts the run first and stamps its id and time', () => {
    vi.useFakeTimers({ now: 1_700_000_000_000 });
    const old: RecentRun = { ...run(), id: 'r-1', ts: 1 };
    expect(addRun(run({ rowCount: 1 }))([old])).toEqual([
      { ...run({ rowCount: 1 }), id: 'r-1700000000000', ts: 1_700_000_000_000 },
      old,
    ]);
  });

  it('an explicit undefined id or time still gets stamped', () => {
    const r = addRun({ ...run(), id: undefined, ts: undefined })([])[0];
    expect(r.id).toMatch(/^r-\d+$/);
    expect(typeof r.ts).toBe('number');
  });

  it('keeps an id the mode chose (Retrieve Multiple updates its run later)', () => {
    expect(addRun(run({ id: 'mine' }))([])[0].id).toBe('mine');
  });

  it(`keeps the list at RECENT_SHOWN (${RECENT_SHOWN})`, () => {
    const many = Array.from({ length: RECENT_SHOWN }, (_, i) => ({
      ...run(),
      id: `r-${i}`,
      ts: i,
    }));
    const next = addRun(run({ id: 'new' }))(many);
    expect(next).toHaveLength(RECENT_SHOWN);
    expect(next[0].id).toBe('new');
    expect(next.map((r) => r.id)).not.toContain(`r-${RECENT_SHOWN - 1}`);
  });
});
