import { describe, expect, it } from 'vitest';
import { noStatusAdvice } from './noStatusAdvice';

describe('noStatusAdvice (#209)', () => {
  it('names Dataverse when the message carries a Dataverse code', () => {
    expect(
      noStatusAdvice({ message: 'Dataverse update failed: 0x80060882: The version …' }),
    ).toMatch(/comes from Dataverse/);
    expect(noStatusAdvice({ code: '0x80040217', message: 'Not found' })).toMatch(
      /comes from Dataverse/,
    );
  });

  it.each([
    [
      'the response-size limit',
      'Dataverse queryData failed: Dataverse response exceeds the maximum size (16 MB)',
    ],
    ['a network failure', 'Request failed: getaddrinfo ENOTFOUND org.crm.dynamics.com'],
    ['an empty query', 'odataQuery parameter cannot be empty'],
  ])('names the host for its own message: %s', (_, message) => {
    expect(noStatusAdvice({ message })).toMatch(/comes from the host itself, not from Dataverse/);
  });

  it('claims neither source otherwise (e.g. a Dataverse error with an empty code)', () => {
    const advice = noStatusAdvice({ message: "Resource not found for the segment 'acounts'." });
    expect(advice).not.toMatch(/comes from/);
    expect(noStatusAdvice(undefined)).not.toMatch(/comes from/);
  });

  it("doesn't take any hex value for a Dataverse code", () => {
    expect(
      noStatusAdvice({ message: 'Cannot create a string longer than 0x1fffffe8 characters' }),
    ).not.toMatch(/comes from Dataverse/);
  });
});
