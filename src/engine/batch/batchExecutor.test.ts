import { describe, expect, it } from 'vitest';
import { compileBatch, newGroup, newItem, planDispatch } from './batchModel';
import {
  describeHostError,
  mapSegmentResults,
  summarizeBatchRun,
  toHostRequest,
  type BatchItemResult,
  type BatchRunResult,
  type SegmentResult,
} from './batchExecutor';

function segmentOf(atomic: boolean, n: number) {
  const items = Array.from({ length: n }, (_, i) =>
    newItem('delete', { entitySet: 'accounts', recordId: String(i + 1) }),
  );
  return planDispatch(compileBatch([newGroup(atomic, items)]).items)[0];
}

describe('mapSegmentResults', () => {
  it('matches results by Content-ID', () => {
    const seg = segmentOf(true, 2);
    const { items, segment } = mapSegmentResults(seg, [
      { contentId: '2', status: 204, headers: {} },
      { contentId: '1', status: 204, headers: {} },
    ]);
    expect(items.map((i) => [i.contentId, i.outcome, i.status])).toEqual([
      ['1', 'done', 204],
      ['2', 'done', 204],
    ]);
    expect(segment.ok).toBe(true);
  });

  it('matches by position when no Content-IDs come back', () => {
    const seg = segmentOf(false, 2);
    const { items } = mapSegmentResults(seg, [
      { status: 204, headers: {} },
      { status: 404, headers: {}, body: { error: { message: 'gone' } } },
    ]);
    expect(items.map((i) => i.status)).toEqual([204, 404]);
  });

  it('marks a failed change set as rolled back and keeps the orphan error', () => {
    const seg = segmentOf(true, 3);
    const error = { status: 400, headers: {}, body: { error: { message: 'bad' } } };
    const { items, segment } = mapSegmentResults(seg, [error]);
    expect(items.every((i) => i.outcome === 'rolled-back')).toBe(true);
    expect(segment.ok).toBe(false);
    expect(segment.changesetError).toBe(error);
  });

  it('marks requests after a stop as not run', () => {
    const seg = segmentOf(false, 3);
    const { items, segment } = mapSegmentResults(seg, [
      { contentId: '1', status: 204, headers: {} },
      { contentId: '2', status: 400, headers: {} },
    ]);
    expect(items.map((i) => i.outcome)).toEqual(['done', 'done', 'not-run']);
    expect(segment.ok).toBe(false);
  });
});

describe('summarizeBatchRun', () => {
  const seg = (s: Partial<SegmentResult> = {}): SegmentResult => ({
    transaction: false,
    contentIds: [],
    ms: 0,
    ok: true,
    ...s,
  });
  const item = (outcome: BatchItemResult['outcome'], status?: number): BatchItemResult => ({
    itemId: 'i',
    contentId: '1',
    outcome,
    status,
  });

  it('records the batch status, not a failing part (#204)', () => {
    // Change set (204) + independent group (200, 404).
    const res: BatchRunResult = {
      ms: 10,
      ok: false,
      segments: [seg({ transaction: true }), seg({ ok: false })],
      items: [item('done', 204), item('done', 200), item('done', 404)],
    };
    expect(summarizeBatchRun(res)).toEqual({
      status: 200,
      summary: '3 requests · 1 failed',
      succeeded: 2,
    });
  });

  it('says when every part succeeded', () => {
    const res: BatchRunResult = {
      ms: 1,
      ok: true,
      segments: [seg()],
      items: [item('done', 204)],
    };
    expect(summarizeBatchRun(res).summary).toBe('1 request · all succeeded');
  });

  it('counts rolled-back and not-run parts', () => {
    const res: BatchRunResult = {
      ms: 1,
      ok: false,
      segments: [seg({ transaction: true, ok: false }), seg({ ok: false, skipped: true })],
      items: [item('done', 400), item('rolled-back'), item('not-run')],
    };
    expect(summarizeBatchRun(res)).toMatchObject({
      status: 200,
      summary: '3 requests · 1 failed · 1 rolled back · 1 not run',
    });
  });

  it('has no status when Dataverse never answered', () => {
    const res: BatchRunResult = {
      ms: 1,
      ok: false,
      segments: [seg({ ok: false, error: 'header consent declined' })],
      items: [item('unknown'), item('unknown')],
    };
    expect(summarizeBatchRun(res)).toMatchObject({
      status: 0,
      summary: '2 requests · 2 no result',
    });
  });

  it('treats a partial response the host could not parse as answered', () => {
    const res: BatchRunResult = {
      ms: 1,
      ok: false,
      segments: [seg({ ok: false, error: 'missing operation results', requestFailed: true })],
      items: [item('unknown')],
    };
    expect(summarizeBatchRun(res).status).toBe(200);
  });
});

describe('toHostRequest', () => {
  it('passes the canonical body text and drops empty headers', () => {
    const c = compileBatch([
      newGroup(true, [newItem('create', { entitySet: 'accounts', body: '{"name":"x"}' })]),
    ]).items[0];
    expect(toHostRequest(c)).toEqual({
      method: 'POST',
      url: 'accounts',
      contentId: '1',
      body: '{\r\n  "name": "x"\r\n}',
    });
  });
});

describe('describeHostError', () => {
  const ctx = { transaction: false, continueOnError: false, sentHeaders: true };
  it('explains a declined consent prompt', () => {
    expect(describeHostError('Additional headers were rejected by the user', ctx)).toMatch(
      /weren’t approved/,
    );
  });
  it('explains the lost partial results without continue-on-error', () => {
    expect(
      describeHostError(
        'Dataverse executeBatch failed: Malformed Dataverse batch response: invalid content ID correlation',
        ctx,
      ),
    ).toMatch(/Continue on error/);
    expect(
      describeHostError('Malformed Dataverse batch response: missing operation results', {
        ...ctx,
        continueOnError: true,
      }),
    ).toBe('Malformed Dataverse batch response: missing operation results');
  });
  it('explains the host parser bug with continue-on-error on and strips the IPC prefix (#198)', () => {
    const msg = describeHostError(
      "Error invoking remote method 'dataverse.executeBatch': Error: Dataverse executeBatch failed: Malformed Dataverse batch response: invalid content ID correlation",
      { ...ctx, continueOnError: true },
    );
    expect(msg).toMatch(/put requests that might fail last/);
    expect(msg).toMatch(
      /Host message: Dataverse executeBatch failed: Malformed Dataverse batch response/,
    );
    expect(msg).not.toMatch(/Error invoking remote method/);
  });
});
