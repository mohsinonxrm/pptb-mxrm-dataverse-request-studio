import { describe, expect, it } from 'vitest';
import {
  MAX_WIDTH_TABLES,
  widthTableKey,
  widthsForTable,
  withColumnWidths,
  withoutTableWidths,
} from './columnWidths';

describe('widthTableKey', () => {
  it('uses the lower-cased logical name, `_` without a table', () => {
    expect(widthTableKey('Account')).toBe('account');
    expect(widthTableKey(' contact ')).toBe('contact');
    expect(widthTableKey(undefined)).toBe('_');
    expect(widthTableKey('')).toBe('_');
  });
});

describe('withColumnWidths', () => {
  it('stores widths per table and grid column id', () => {
    let store = withColumnWidths(undefined, 'account', { name: 250.4, _ownerid_value: 180 });
    store = withColumnWidths(store, 'contact', { name: 300 });
    store = withColumnWidths(store, 'Account', { 'primarycontactid.fullname': 222, name: 260 });
    expect(widthsForTable(store, 'account')).toEqual({
      name: 260,
      _ownerid_value: 180,
      'primarycontactid.fullname': 222,
    });
    expect(widthsForTable(store, 'contact')).toEqual({ name: 300 });
    expect(widthsForTable(store, 'lead')).toEqual({});
  });

  it('keeps raw twin columns separate from their formatted column', () => {
    const store = withColumnWidths({}, 'account', { statuscode: 140, statuscode__raw: 90 });
    expect(widthsForTable(store, 'account')).toEqual({ statuscode: 140, statuscode__raw: 90 });
  });

  it('drops nonsense widths', () => {
    const store = withColumnWidths({}, 'account', { a: NaN, b: 5, c: Infinity, d: 100 });
    expect(widthsForTable(store, 'account')).toEqual({ d: 100 });
  });

  it('keeps only the most recently resized tables', () => {
    let store = {};
    for (let i = 0; i < MAX_WIDTH_TABLES + 5; i++)
      store = withColumnWidths(store, `t${i}`, { a: 100 });
    // Touch an old-but-kept table so it becomes most recent.
    store = withColumnWidths(store, 't10', { a: 120 });
    store = withColumnWidths(store, 'new', { a: 100 });
    const keys = Object.keys(store);
    expect(keys).toHaveLength(MAX_WIDTH_TABLES);
    expect(keys.slice(-2)).toEqual(['t10', 'new']);
    expect(keys).not.toContain('t5');
  });

  it('does not mutate the input store', () => {
    const store = { account: { name: 100 } };
    withColumnWidths(store, 'account', { name: 200 });
    expect(store).toEqual({ account: { name: 100 } });
  });
});

describe('withoutTableWidths', () => {
  it('removes just that table', () => {
    const store = { account: { name: 100 }, contact: { name: 200 } };
    expect(withoutTableWidths(store, 'Account')).toEqual({ contact: { name: 200 } });
    expect(withoutTableWidths(undefined, 'account')).toEqual({});
  });
});
