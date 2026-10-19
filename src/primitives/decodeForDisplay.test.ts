import { describe, expect, it } from 'vitest';
import { decodeForDisplay } from './decodeForDisplay';

describe('decodeForDisplay (#48)', () => {
  it('decodes percent-encoded values', () => {
    expect(decodeForDisplay('caf%C3%A9%20%26%20bar')).toBe('café & bar');
  });

  it('makes a double-encoded $skiptoken cookie readable', () => {
    // URLSearchParams has already decoded once; this is the second pass.
    expect(decodeForDisplay('%3Ccookie pagenumber=%222%22 /%3E')).toBe('<cookie pagenumber="2" />');
  });

  it('returns the value unchanged when it holds a literal % (no URIError)', () => {
    expect(() => decodeForDisplay("contains(name,'50%')")).not.toThrow();
    expect(decodeForDisplay("contains(name,'50%')")).toBe("contains(name,'50%')");
    expect(decodeForDisplay('100%ZZ')).toBe('100%ZZ');
  });

  it('returns the value unchanged for an invalid UTF-8 escape', () => {
    expect(decodeForDisplay('%E9')).toBe('%E9');
  });

  it('leaves plain values alone', () => {
    expect(decodeForDisplay('')).toBe('');
    expect(decodeForDisplay("name eq 'x'")).toBe("name eq 'x'");
  });
});
