// #292: the Columns menu's per-table hidden columns.
import { describe, expect, it } from 'vitest';
import {
  MAX_HIDDEN_TABLES,
  columnsButtonLabel,
  hiddenForTable,
  visibleColumnKeys,
  withHiddenColumns,
} from './hiddenColumns';

describe('hidden columns', () => {
  it('stores per table, case-insensitive, and Show all removes the table', () => {
    let s = withHiddenColumns({}, 'Account', ['name', 'name', 'revenue']);
    expect(hiddenForTable(s, 'account')).toEqual(['name', 'revenue']);
    expect(hiddenForTable(s, 'contact')).toEqual([]);
    s = withHiddenColumns(s, 'account', []);
    expect(s).toEqual({});
  });

  it('keeps the most recent tables only', () => {
    let s = {};
    for (let i = 0; i <= MAX_HIDDEN_TABLES; i++) s = withHiddenColumns(s, `t${i}`, ['x']);
    expect(Object.keys(s)).toHaveLength(MAX_HIDDEN_TABLES);
    expect(hiddenForTable(s, 't0')).toEqual([]);
    expect(hiddenForTable(s, `t${MAX_HIDDEN_TABLES}`)).toEqual(['x']);
  });

  it('shows what is not hidden, and everything rather than nothing', () => {
    expect(visibleColumnKeys(['a', 'b', 'c'], ['b'])).toEqual(['a', 'c']);
    expect(visibleColumnKeys(['a', 'b'], ['a', 'b'])).toEqual(['a', 'b']);
    expect(hiddenForTable({ x: 'junk' } as never, 'x')).toEqual([]);
  });
});

describe('Columns button label', () => {
  it('counts shown of all only when some are hidden', () => {
    const c = (shown: boolean) => ({ key: 'k', label: 'K', shown });
    expect(columnsButtonLabel([c(true), c(true)])).toBe('Columns');
    expect(columnsButtonLabel([c(true), c(false), c(false)])).toBe('Columns · 1 of 3');
  });
});
