import { describe, expect, it } from 'vitest';
import { jsonBytes, utf8ByteLength } from './byteSize';

const encoded = (s: string) => new TextEncoder().encode(s).length;

describe('utf8ByteLength', () => {
  it.each([
    ['ascii', 'Contoso'],
    ['2-byte', 'Müller'],
    ['3-byte', '東京'],
    ['surrogate pair', 'rocket 🚀'],
    ['lone high surrogate', 'a\ud800b'],
    ['lone low surrogate', 'a\udc00b'],
    ['empty', ''],
  ])('matches TextEncoder for %s text', (_, s) => {
    expect(utf8ByteLength(s)).toBe(encoded(s));
  });
});

describe('jsonBytes', () => {
  it('counts UTF-8 bytes, not UTF-16 code units (#223)', () => {
    const body = { value: [{ name: '東京 🚀' }] };
    expect(JSON.stringify(body).length).toBeLessThan(jsonBytes(body));
    expect(jsonBytes(body)).toBe(encoded(JSON.stringify(body)));
  });

  it('is 0 for no body', () => {
    expect(jsonBytes(null)).toBe(0);
    expect(jsonBytes(undefined)).toBe(0);
  });
});
