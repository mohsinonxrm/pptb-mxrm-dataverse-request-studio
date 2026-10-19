import { describe, expect, it } from 'vitest';
import { buildSearchIndex, cellSearchText, filterRows, isFiltering } from './gridFilter';

const FV = '@OData.Community.Display.V1.FormattedValue';

const rows = [
  {
    name: 'Contoso',
    statuscode: 1,
    [`statuscode${FV}`]: 'Active',
    _ownerid_value: 'aaaa-1111',
    [`_ownerid_value${FV}`]: 'Jane Doe',
    hidden: 'secret',
  },
  {
    name: 'Fabrikam',
    statuscode: 2,
    [`statuscode${FV}`]: 'Inactive',
    _ownerid_value: 'bbbb-2222',
    [`_ownerid_value${FV}`]: 'John Roe',
    hidden: 'other',
  },
];
const keys = ['name', 'statuscode', '_ownerid_value'];
const index = buildSearchIndex(rows, keys);

describe('cellSearchText', () => {
  it('covers formatted and raw values, lower-cased', () => {
    expect(cellSearchText(rows[0], 'statuscode')).toContain('active');
    expect(cellSearchText(rows[0], 'statuscode')).toContain('1');
    expect(cellSearchText(rows[0], '_ownerid_value')).toContain('jane doe');
  });
  it('uses only the raw value for a (Raw) twin column', () => {
    expect(cellSearchText(rows[0], 'statuscode__raw')).toBe('1');
  });
});

describe('filterRows', () => {
  it('returns the same array when nothing filters', () => {
    expect(filterRows(rows, index, '  ', { name: ' ' })).toBe(rows);
    expect(isFiltering('  ', { name: ' ' })).toBe(false);
  });

  it('searches visible columns, formatted or raw, case-insensitively', () => {
    expect(filterRows(rows, index, 'JANE', {})).toEqual([rows[0]]);
    expect(filterRows(rows, index, 'bbbb', {})).toEqual([rows[1]]);
    expect(filterRows(rows, index, 'active', {})).toEqual(rows); // "Inactive" contains it too
    // Columns the grid doesn't show aren't searched.
    expect(filterRows(rows, index, 'secret', {})).toEqual([]);
  });

  it('does not match across cell boundaries', () => {
    expect(filterRows(rows, index, 'contoso1', {})).toEqual([]);
  });

  it('applies every column filter together with the search', () => {
    expect(filterRows(rows, index, '', { statuscode: 'inactive' })).toEqual([rows[1]]);
    expect(filterRows(rows, index, '', { name: 'o', _ownerid_value: 'jane' })).toEqual([rows[0]]);
    expect(filterRows(rows, index, 'fabrikam', { _ownerid_value: 'jane' })).toEqual([]);
    expect(filterRows(rows, index, '', { statuscode__raw: '2' })).toEqual([rows[1]]);
    expect(isFiltering('', { statuscode: 'x' })).toBe(true);
  });
});
