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

describe('ETag paste mistakes (#264 review)', () => {
  it.each([['W/""'], ['""'], ['W/"1 2"']])('rejects an empty or spaced tag: %s', (v) => {
    expect(etagProblem(v)).not.toBeNull();
  });

  it('explains a JSON-escaped ETag and fixes it', () => {
    expect(etagProblem('W/\\"7281965\\"')).toMatch(/JSON-escaped/);
    expect(etagFix('W/\\"7281965\\"')).toBe('W/"7281965"');
    // The whole JSON string value, outer quotes included.
    expect(etagFix('"W/\\"7281965\\""')).toBe('W/"7281965"');
  });

  it.each([
    ['lowercase w/', 'w/"7281965"', 'W/"7281965"'],
    ['a space after W/', 'W/ "7281965"', 'W/"7281965"'],
    ['single curly quotes', 'W/‘7281965’', 'W/"7281965"'],
    ['fullwidth quotes', 'W/＂7281965＂', 'W/"7281965"'],
  ])('fixes %s', (_, v, want) => {
    expect(etagProblem(v)).not.toBeNull();
    expect(etagFix(v)).toBe(want);
  });

  it('flags curly quotes the same way every time (no stateful regex)', () => {
    expect(etagProblem('W/“1”')).toMatch(/curly/);
    expect(etagProblem('W/“1”')).toMatch(/curly/);
  });

  it('offers no fix for an empty value', () => {
    expect(etagFix('')).toBeNull();
  });
});
