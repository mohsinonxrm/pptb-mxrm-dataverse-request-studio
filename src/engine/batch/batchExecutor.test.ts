import { describe, expect, it } from 'vitest';
import { compileBatch, newGroup, newItem, planDispatch } from './batchModel';
import { describeHostError, mapSegmentResults, toHostRequest } from './batchExecutor';

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
});
