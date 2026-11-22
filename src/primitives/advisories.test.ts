import { describe, expect, it } from 'vitest';
import { disabledReasonFromAdvisories, firstBlocker } from './advisories';

describe('firstBlocker (#294)', () => {
  it('reports only the first failing check, as an error advisory with its node', () => {
    const out = firstBlocker([
      { when: false, id: 'a', title: 'A', focusNode: 'target' },
      { when: 'x', id: 'b', title: 'B', focusNode: 'fieldset' },
      { when: true, id: 'c', title: 'C', focusNode: 'headers' },
    ]);
    expect(out).toEqual([
      {
        id: 'b',
        severity: 'error',
        source: 'validation',
        title: 'B',
        body: undefined,
        focusNode: 'fieldset',
        learnMoreUrl: undefined,
      },
    ]);
    // The Execute tooltip stays the blocker's own message.
    expect(disabledReasonFromAdvisories(out)).toBe('B');
  });

  it('is empty when nothing blocks', () => {
    expect(firstBlocker([{ when: null, id: 'a', title: 'A' }])).toEqual([]);
    expect(firstBlocker([])).toEqual([]);
  });
});
