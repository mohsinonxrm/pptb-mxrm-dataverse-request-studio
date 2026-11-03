// #387: every results row's checkbox has a name.
import { describe, expect, it } from 'vitest';
import { rowSelectLabel } from './rowSelectLabel';

describe('rowSelectLabel', () => {
  it('names the row after its primary name', () => {
    expect(rowSelectLabel({ fullname: ' Ada Lovelace ' }, 'fullname', 0)).toBe(
      'Select Ada Lovelace',
    );
  });

  it('falls back to the row number without a primary name value', () => {
    expect(rowSelectLabel({ fullname: '' }, 'fullname', 2)).toBe('Select row 3');
    expect(rowSelectLabel({ fullname: null }, 'fullname', 0)).toBe('Select row 1');
    expect(rowSelectLabel({ count: 4 }, undefined, 9)).toBe('Select row 10');
  });
});
