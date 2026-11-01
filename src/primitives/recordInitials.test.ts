import { describe, expect, it } from 'vitest';
import { recordInitials } from './recordInitials';

describe('recordInitials (#224)', () => {
  it.each([
    ['Contoso Ltd', 'CL'],
    ['  adventure   works  cycles ', 'AC'],
    ['Fabrikam', 'F'],
    ['émile zola', 'ÉZ'],
    // Decomposed: e + combining acute.
    ['émile zola', 'ÉZ'],
    ['(Test) Account', 'TA'],
    ['#1 Supplier', '1S'],
    ['😀 smile', 'S'],
    ['🇺🇸 👨‍👩‍👧', '?'],
    ['ßeta', 'ß'],
    ['', '?'],
    ['   ', '?'],
  ])('%j → %s', (name, initials) => {
    expect(recordInitials(name)).toBe(initials);
  });
});
