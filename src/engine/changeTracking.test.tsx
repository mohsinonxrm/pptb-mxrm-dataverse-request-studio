// Change tracking in Retrieve Multiple (#233): when the option is offered,
// which query options block Execute, and reading a delta response.
import { beforeEach, describe, expect, it } from 'vitest';
import {
  CHANGE_KEY,
  DELETED_ID_KEY,
  changeTrackingAdvisories,
  changeTrackingAvailability,
  deletedRowCount,
  deltaGridRow,
  deltaLinkOf,
  isDeletedEntry,
} from './changeTracking';
import { buildRetrieveMultiple } from './urlBuilder';
import {
  account,
  group,
  registerFixtureTables,
  retrieveMultipleState,
  rule,
} from '../test/fixtures';
import { emptyPrefer } from '../editors/PreferEditor';
import { makeExpandSpec } from '../editors/ExpandEditor';
import { disabledReasonFromAdvisories } from '../primitives/advisories';
import type { RetrieveMultipleState } from '../state/readState';
import type { TableMeta } from '../metadata/metadata';

const tracked: TableMeta = { ...account, changeTrackingEnabled: true };

beforeEach(() => registerFixtureTables());

describe('changeTrackingAvailability', () => {
  it('is offered for a table with change tracking enabled', () => {
    expect(changeTrackingAvailability({ table: 'account', tableMeta: tracked })).toEqual({
      available: true,
    });
  });

  it("is off, with the reason, when ChangeTrackingEnabled is false or isn't there", () => {
    const reason =
      "Change tracking isn't enabled for this table. Turn on Track changes in the table's " +
      'properties in Power Apps (Advanced options) to use it.';
    expect(
      changeTrackingAvailability({
        table: 'account',
        tableMeta: { ...account, changeTrackingEnabled: false },
      }),
    ).toEqual({ available: false, reason });
    expect(changeTrackingAvailability({ table: 'account', tableMeta: account })).toEqual({
      available: false,
      reason,
    });
  });

  it('is off without a table, while it loads, and with Count only', () => {
    expect(changeTrackingAvailability({ table: '', tableMeta: undefined }).reason).toBe(
      'Pick a target table first.',
    );
    expect(changeTrackingAvailability({ table: 'account', tableMeta: undefined }).reason).toBe(
      "Loading the table's definition…",
    );
    expect(
      changeTrackingAvailability({ table: 'account', tableMeta: tracked, countOnly: true }),
    ).toEqual({
      available: false,
      reason: "Count only doesn't track changes. Turn Count only off in $count to use it.",
    });
  });
});

describe('changeTrackingAdvisories', () => {
  const advisoriesFor = (o: Partial<RetrieveMultipleState>, meta = tracked) => {
    const s = retrieveMultipleState({
      prefer: { ...emptyPrefer(), trackChanges: true },
      ...o,
    });
    return changeTrackingAdvisories({
      trackChanges: s.prefer.trackChanges,
      countOnly: s.countOnly,
      tableMeta: meta,
      queryParts: buildRetrieveMultiple(s).queryParts,
    });
  };

  it('allows $select', () => {
    expect(advisoriesFor({ select: ['name'] })).toEqual([]);
  });

  it('blocks $filter, $orderby, $expand and $top, each with its own advisory', () => {
    const list = advisoriesFor({
      select: ['name'],
      filter: group('and', [rule('name', 'eq', 'Contoso')]),
      orderby: [{ col: 'name', dir: 'asc' }],
      expand: [makeExpandSpec('primarycontactid')],
      top: 50,
    });
    expect(list.map((a) => [a.severity, a.title, a.focusNode])).toEqual([
      ['error', "Change tracking can't be combined with $filter", 'filter'],
      ['error', "Change tracking can't be combined with $orderby", 'orderby'],
      ['error', "Change tracking can't be combined with $expand", 'expand'],
      ['error', "Change tracking can't be combined with $top", 'top'],
    ]);
    expect(disabledReasonFromAdvisories(list)).toBe('4 blockers — open Advisories to review.');
  });

  it("names the one option when it's the only one", () => {
    const list = advisoriesFor({ top: 50 });
    expect(disabledReasonFromAdvisories(list)).toBe("Change tracking can't be combined with $top");
  });

  it("goes by what the URL sends: $top isn't sent next to a page size", () => {
    expect(
      advisoriesFor({ top: 50, prefer: { ...emptyPrefer(), trackChanges: true, maxpagesize: 10 } }),
    ).toEqual([]);
  });

  it("blocks a table that isn't tracked, once its definition is loaded", () => {
    expect(advisoriesFor({}, account).map((a) => [a.title, a.focusNode])).toEqual([
      ["Change tracking isn't enabled for this table", 'prefer'],
    ]);
  });

  it('says nothing when Track changes is off, or with Count only', () => {
    const filtered = { filter: group('and', [rule('name', 'eq', 'x')]), top: 50 };
    expect(advisoriesFor({ ...filtered, prefer: emptyPrefer() })).toEqual([]);
    expect(advisoriesFor({ ...filtered, countOnly: true })).toEqual([]);
  });
});

describe('delta responses', () => {
  const deleted = {
    '@odata.context':
      'https://org.crm.dynamics.com/api/data/v9.2/$metadata#accounts/$deletedEntity',
    id: '2e451703-c686-e711-80e5-00155db19e6d',
    reason: 'deleted',
  };
  const changed = { name: 'Monte Orton', accountid: '60c4e274-0d87-e711-80e5-00155db19e6d' };

  it('reads the delta link', () => {
    expect(
      deltaLinkOf({ '@odata.deltaLink': 'https://x/api/data/v9.2/accounts?$deltatoken=1' }),
    ).toBe('https://x/api/data/v9.2/accounts?$deltatoken=1');
    expect(deltaLinkOf({ value: [] })).toBeUndefined();
    expect(deltaLinkOf(5)).toBeUndefined();
  });

  it('tells a deleted entry from a changed row', () => {
    expect(isDeletedEntry(deleted)).toBe(true);
    expect(isDeletedEntry(changed)).toBe(false);
    expect(isDeletedEntry({ id: 'x', reason: 'deleted' })).toBe(false);
    expect(deletedRowCount({ value: [changed, deleted, deleted] })).toBe(2);
    expect(deletedRowCount({})).toBe(0);
  });

  it('makes a deleted entry a row with the Change column, and leaves others as they are', () => {
    expect(deltaGridRow(deleted)).toEqual({
      [CHANGE_KEY]: 'Deleted 2e451703-c686-e711-80e5-00155db19e6d',
      [DELETED_ID_KEY]: '2e451703-c686-e711-80e5-00155db19e6d',
    });
    expect(deltaGridRow(changed)).toBe(changed);
  });
});
