import { describe, expect, it } from 'vitest';
import { GROUPS, REQ_TYPES } from './requestTypes';

// #289 — every request type explains itself in the picker.
describe('request type descriptions', () => {
  it.each(REQ_TYPES.map((r) => [r.id, r] as const))('%s has a one-sentence description', (_, r) => {
    expect(r.description.trim().length).toBeGreaterThan(15);
    expect(r.description.toLowerCase()).not.toBe(r.name.toLowerCase());
    expect(r.description).toMatch(/\.$/);
  });

  it('every group with request types has a tagline', () => {
    for (const g of GROUPS) {
      if (REQ_TYPES.some((r) => r.group === g.id)) expect(g.tagline.trim()).not.toBe('');
    }
  });
});
