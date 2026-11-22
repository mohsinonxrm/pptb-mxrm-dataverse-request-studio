// useTableDetail: the JSON body editor's schema needs every column's
// options, lookup targets and date behavior, which the table's first load
// leaves out (#268). One type-cast request per subtype fills them.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { useTableDetail } from './useColumnDetail';
import {
  __clearLiveTables,
  __registerLiveTable,
  findTable,
  type ColumnMeta,
  type TableMeta,
} from '../metadata/metadata';
import { account } from '../test/fixtures';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const org = vi.hoisted(() => ({ current: 'a.crm.dynamics.com' }));
vi.mock('./activeOrg', () => ({
  activeOrg: () => org.current,
  onActiveOrgChange: () => () => {},
}));

const loadAttributesByTypeCast = vi.hoisted(() => vi.fn());
vi.mock('./dataverseMetadata', async (importOriginal) => ({
  ...(await importOriginal<typeof import('./dataverseMetadata')>()),
  loadAttributesByTypeCast,
}));

// Account as the first load returns it: no options, targets or labels. New
// column objects each time, as a metadata refresh builds.
const bare = (): TableMeta => ({
  ...account,
  columns: account.columns.map((c): ColumnMeta => {
    if (c.attributeType === 'Picklist' || c.attributeType === 'MultiSelectPicklist')
      return { ...c, options: [] };
    if (c.attributeType === 'Lookup') return { ...c, targets: [] };
    if (c.attributeType === 'Boolean')
      return {
        ...c,
        trueOption: { value: 1, label: 'Yes' },
        falseOption: { value: 0, label: 'No' },
      };
    return { ...c };
  }),
});

const rowsFor: Record<string, unknown[]> = {
  'Microsoft.Dynamics.CRM.PicklistAttributeMetadata': [
    {
      LogicalName: 'industrycode',
      DefaultFormValue: -1,
      OptionSet: {
        Options: [{ Value: 1, Label: { UserLocalizedLabel: { Label: 'Accounting' } } }],
      },
    },
  ],
  'Microsoft.Dynamics.CRM.MultiSelectPicklistAttributeMetadata': [],
  'Microsoft.Dynamics.CRM.BooleanAttributeMetadata': [
    {
      LogicalName: 'donotemail',
      OptionSet: {
        TrueOption: { Value: 1, Label: { UserLocalizedLabel: { Label: 'Do Not Allow' } } },
        FalseOption: { Value: 0, Label: { UserLocalizedLabel: { Label: 'Allow' } } },
      },
    },
  ],
  'Microsoft.Dynamics.CRM.LookupAttributeMetadata': [
    { LogicalName: 'primarycontactid', Targets: ['contact'] },
  ],
  'Microsoft.Dynamics.CRM.DateTimeAttributeMetadata': [
    { LogicalName: 'createdon', Format: 'DateAndTime', DateTimeBehavior: { Value: 'UserLocal' } },
  ],
};

let root: Root;
let host: HTMLDivElement;
const seen: { loading: boolean; table?: TableMeta }[] = [];
function Probe({ name }: { name: string }) {
  seen.push(useTableDetail(name));
  return null;
}
const render = async (name = 'account') => {
  await act(async () => root.render(<Probe name={name} />));
  // Let the fetches land and the patched table re-render
  await act(async () => {
    await new Promise((r) => setTimeout(r, 0));
  });
};
const col = (n: string) => findTable('account')!.columns.find((c) => c.logicalName === n)!;
// A registry update for the table that keeps its column objects
const touch = async () => {
  await act(async () => {
    const t = findTable('account')!;
    __registerLiveTable({ ...t, columns: [...t.columns] });
  });
  await act(async () => {
    await new Promise((r) => setTimeout(r, 0));
  });
};
const last = () => seen[seen.length - 1];

beforeEach(() => {
  org.current = 'a.crm.dynamics.com';
  __clearLiveTables();
  __registerLiveTable(bare());
  seen.length = 0;
  loadAttributesByTypeCast.mockReset();
  loadAttributesByTypeCast.mockImplementation(async (_t: string, cast: string) => rowsFor[cast]);
  host = document.createElement('div');
  root = createRoot(host);
});
afterEach(() => {
  act(() => root.unmount());
  __clearLiveTables();
});

describe('useTableDetail', () => {
  it('fills options, targets and labels with one request per subtype', async () => {
    await render();
    expect(seen[0].loading).toBe(true);
    expect(last().loading).toBe(false);
    expect(loadAttributesByTypeCast.mock.calls.map((c) => c[1]).sort()).toEqual(
      [
        'Microsoft.Dynamics.CRM.BooleanAttributeMetadata',
        'Microsoft.Dynamics.CRM.DateTimeAttributeMetadata',
        'Microsoft.Dynamics.CRM.LookupAttributeMetadata',
        'Microsoft.Dynamics.CRM.MultiSelectPicklistAttributeMetadata',
        'Microsoft.Dynamics.CRM.PicklistAttributeMetadata',
      ].sort(),
    );
    expect(col('industrycode')).toMatchObject({ options: [{ value: 1, label: 'Accounting' }] });
    expect(col('primarycontactid')).toMatchObject({ targets: ['contact'] });
    expect(col('donotemail')).toMatchObject({ trueOption: { label: 'Do Not Allow' } });
    expect(last().table).toBe(findTable('account'));
  });

  it('doesn’t fetch again for columns the response left as they were', async () => {
    await render();
    const calls = loadAttributesByTypeCast.mock.calls.length;
    // createdon stays UserLocal, sample_tags has no row: neither counts as
    // enriched, and the table changing doesn't fetch them again
    await touch();
    expect(loadAttributesByTypeCast.mock.calls.length).toBe(calls);
    expect(last().loading).toBe(false);
  });

  it('stops loading when the requests fail, and tries again when the table changes', async () => {
    loadAttributesByTypeCast.mockRejectedValue(new Error('offline'));
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    await render();
    expect(last().loading).toBe(false);
    expect(col('industrycode')).toMatchObject({ options: [] });
    expect(warn).toHaveBeenCalled();
    warn.mockRestore();

    loadAttributesByTypeCast.mockImplementation(async (_t: string, cast: string) => rowsFor[cast]);
    await touch();
    expect(col('industrycode')).toMatchObject({ options: [{ value: 1, label: 'Accounting' }] });
    expect(last().loading).toBe(false);
  });

  it('drops a response that lands after a connection switch, and fetches again back there', async () => {
    const pending: (() => void)[] = [];
    loadAttributesByTypeCast.mockImplementation(
      (_t: string, cast: string) =>
        new Promise((resolve) => pending.push(() => resolve(rowsFor[cast]))),
    );
    await render();
    expect(last().loading).toBe(true);
    org.current = 'b.crm.dynamics.com';
    __registerLiveTable(bare());
    await act(async () => {
      for (const p of pending) p();
      await new Promise((r) => setTimeout(r, 0));
    });
    // Org b's table isn't patched with org a's rows
    expect(col('industrycode')).toMatchObject({ options: [] });

    org.current = 'a.crm.dynamics.com';
    const calls = loadAttributesByTypeCast.mock.calls.length;
    loadAttributesByTypeCast.mockImplementation(async (_t: string, cast: string) => rowsFor[cast]);
    await touch();
    expect(loadAttributesByTypeCast.mock.calls.length).toBeGreaterThan(calls);
    expect(col('industrycode')).toMatchObject({ options: [{ value: 1, label: 'Accounting' }] });
    expect(last().loading).toBe(false);
  });

  it('fetches only the subtypes still missing', async () => {
    // The fixture's options, targets and labels are loaded; createdon's
    // User Local behavior is also the default, so it's checked once
    __registerLiveTable({ ...account, columns: account.columns.map((c) => ({ ...c })) });
    await render();
    expect(loadAttributesByTypeCast.mock.calls.map((c) => c[1])).toEqual([
      'Microsoft.Dynamics.CRM.DateTimeAttributeMetadata',
    ]);
    expect(last().loading).toBe(false);
  });
});
