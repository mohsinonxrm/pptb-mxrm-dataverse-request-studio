import { describe, expect, it } from 'vitest';
import { noStatusAdvice } from './ResultsView';

describe('noStatusAdvice (#209)', () => {
  it('names Dataverse when the message carries a Dataverse code', () => {
    expect(
      noStatusAdvice({ message: 'Dataverse update failed: 0x80060882: The version …' }),
    ).toMatch(/comes from Dataverse/);
    expect(noStatusAdvice({ code: '0x80040217', message: 'Not found' })).toMatch(
      /comes from Dataverse/,
    );
  });

  it('names the host for its own messages, e.g. the response-size limit', () => {
    expect(
      noStatusAdvice({
        message: 'Dataverse queryData failed: Dataverse response exceeds the maximum size (16 MB)',
      }),
    ).toMatch(/comes from the host itself, not from Dataverse/);
    expect(noStatusAdvice(undefined)).toMatch(/host itself/);
  });
});
