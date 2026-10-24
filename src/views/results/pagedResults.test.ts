import { describe, expect, it, vi } from 'vitest';
import type { ExecResult } from '../../engine/dataverseExecutor';
import { mergePage, retrieveAll, type PageBatch, type PagingStatus } from './pagedResults';

type Row = Record<string, unknown>;

const last = <T>(a: T[]): T => a[a.length - 1];

const rows = (from: number, n: number): Row[] =>
  Array.from({ length: n }, (_, i) => ({ accountid: `a${from + i}` }));

function page(value: Row[], nextLink?: string, bytes = 100): ExecResult {
  return {
    status: 200,
    statusText: 'OK',
    ms: 1,
    bytes,
    ok: true,
    outcome: 'ok',
    body: { '@odata.context': 'ctx', value, ...(nextLink ? { '@odata.nextLink': nextLink } : {}) },
    headers: {},
  } as ExecResult;
}

const failed = { status: 500, ok: false, bytes: 0, body: { error: { message: 'boom' } } };

/** `count` pages of `size` rows, linked p1 → p2 → … */
function server(count: number, size: number) {
  const fetchPage = vi.fn(async (url: string) => {
    const n = Number(url.slice(1));
    return page(rows((n - 1) * size, size), n < count ? `p${n + 1}` : undefined);
  });
  return fetchPage;
}

function collect() {
  const batches: PageBatch[] = [];
  return { batches, commit: (b: PageBatch) => batches.push(b) };
}

describe('mergePage', () => {
  it('appends rows, keeps the envelope, replaces the nextLink and adds the bytes', () => {
    const merged = mergePage(page(rows(0, 2), 'p2', 300), {
      rows: rows(2, 2),
      bytes: 200,
      nextLink: 'p3',
    });
    const body = merged.body as {
      value: Row[];
      '@odata.context': string;
      '@odata.nextLink': string;
    };
    expect(body.value.map((r) => r.accountid)).toEqual(['a0', 'a1', 'a2', 'a3']);
    expect(body['@odata.context']).toBe('ctx');
    expect(body['@odata.nextLink']).toBe('p3');
    expect(merged.bytes).toBe(500);
  });

  it('drops the nextLink once the last page is in', () => {
    const merged = mergePage(page([], 'p2'), { rows: [], bytes: 0 });
    expect(merged.body).not.toHaveProperty(['@odata.nextLink']);
  });
});

describe('retrieveAll', () => {
  const current = (): PagingStatus => 'current';

  it('loads every page and ends with no nextLink', async () => {
    const { batches, commit } = collect();
    const out = await retrieveAll({
      nextLink: 'p1',
      fetchPage: server(3, 10),
      commit,
      status: current,
    });
    expect(out).toEqual({ kind: 'done', rows: 30 });
    expect(batches.flatMap((b) => b.rows)).toHaveLength(30);
    expect(last(batches).nextLink).toBeUndefined();
  });

  it('merges in batches, not once per page', async () => {
    const { batches, commit } = collect();
    let t = 0;
    await retrieveAll({
      nextLink: 'p1',
      fetchPage: server(10, 10),
      commit,
      status: current,
      commitMs: 500,
      now: () => (t += 100), // each page "takes" 100 ms
    });
    expect(batches.length).toBeLessThan(10);
    expect(batches.flatMap((b) => b.rows)).toHaveLength(100);
  });

  it('pauses at the row budget, with the nextLink that resumes it (#223)', async () => {
    const { batches, commit } = collect();
    const out = await retrieveAll({
      nextLink: 'p1',
      fetchPage: server(10, 10),
      commit,
      status: current,
      rowBudget: 25,
    });
    expect(out).toEqual({ kind: 'paused', rows: 30 });
    expect(last(batches).nextLink).toBe('p4');
  });

  it('pauses at the page budget for very small pages', async () => {
    const { commit } = collect();
    const out = await retrieveAll({
      nextLink: 'p1',
      fetchPage: server(10, 1),
      commit,
      status: current,
      pageBudget: 4,
    });
    expect(out).toEqual({ kind: 'paused', rows: 4 });
  });

  it('keeps what it loaded and stops at a failed page', async () => {
    const { batches, commit } = collect();
    const fetchPage = vi.fn(async (url: string) =>
      url === 'p3'
        ? (failed as unknown as ExecResult)
        : page(rows(0, 10), url === 'p1' ? 'p2' : 'p3'),
    );
    const out = await retrieveAll({ nextLink: 'p1', fetchPage, commit, status: current });
    expect(out.kind).toBe('failed');
    expect(batches.flatMap((b) => b.rows)).toHaveLength(20);
    expect(last(batches).nextLink).toBe('p3');
  });

  it('on Stop, keeps the merged rows and drops the page that arrived after it', async () => {
    const { batches, commit } = collect();
    let status: PagingStatus = 'current';
    const fetchPage = vi.fn(async (url: string) => {
      if (url === 'p3') status = 'stopped'; // Stop pressed while p3 loads
      const n = Number(url.slice(1));
      return page(rows((n - 1) * 10, 10), `p${n + 1}`);
    });
    const out = await retrieveAll({ nextLink: 'p1', fetchPage, commit, status: () => status });
    expect(out).toEqual({ kind: 'stopped', rows: 20 });
    expect(batches.flatMap((b) => b.rows)).toHaveLength(20);
    // Resuming fetches p3 again, so nothing is lost or duplicated.
    expect(last(batches).nextLink).toBe('p3');
  });

  it('merges nothing once a newer run has started', async () => {
    const { batches, commit } = collect();
    let status: PagingStatus = 'current';
    const fetchPage = vi.fn(async (url: string) => {
      if (url === 'p2') status = 'superseded';
      return page(rows(0, 10), url === 'p1' ? 'p2' : 'p3');
    });
    const out = await retrieveAll({ nextLink: 'p1', fetchPage, commit, status: () => status });
    expect(out).toEqual({ kind: 'superseded' });
    expect(batches).toHaveLength(0);
  });
});
