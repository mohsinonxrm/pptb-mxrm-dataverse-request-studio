import { describe, expect, it } from 'vitest';
import { etagFix, etagProblem, preconditionToHeader } from './PreconditionEditor';

describe('etagProblem (#264)', () => {
  it.each([
    ['weak', 'W/"7281965"'],
    ['strong', '"7281965"'],
    ['padded', '  W/"7281965"\n'],
    ['any version', '*'],
  ])('accepts a %s ETag', (_, v) => {
    expect(etagProblem(v)).toBeNull();
  });

  it('asks for a value when it is empty', () => {
    expect(etagProblem('  ')).toMatch(/Provide an etag value/);
  });

  it('says to wrap a bare version number', () => {
    expect(etagProblem('7281965')).toBe('Wrap the ETag in quotes: W/"7281965".');
  });

  it('flags curly quotes', () => {
    expect(etagProblem('W/“7281965”')).toMatch(/curly quotes/);
  });

  it.each([['W/7281965'], ['W/"7281965'], ['"72"81965"'], ['etag: W/"1"']])(
    'rejects a malformed value: %s',
    (v) => {
      expect(etagProblem(v)).toMatch(/looks like W\/"12345678"/);
    },
  );
});

describe('etagFix', () => {
  it('wraps a bare version number', () => {
    expect(etagFix(' 7281965 ')).toBe('W/"7281965"');
  });

  it('straightens curly quotes', () => {
    expect(etagFix('W/“7281965”')).toBe('W/"7281965"');
  });

  it('offers nothing for a valid or unfixable value', () => {
    expect(etagFix('W/"7281965"')).toBeNull();
    expect(etagFix('W/7281965')).toBeNull();
  });
});

describe('preconditionToHeader', () => {
  it('sends the ETag trimmed', () => {
    expect(preconditionToHeader({ kind: 'etag', etag: '  W/"7281965" ' })).toEqual({
      name: 'If-Match',
      value: 'W/"7281965"',
    });
  });

  it('sends nothing for an empty ETag', () => {
    expect(preconditionToHeader({ kind: 'etag', etag: ' ' })).toBeNull();
  });
});
