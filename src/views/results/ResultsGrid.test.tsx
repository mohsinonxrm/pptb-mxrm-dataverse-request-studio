// ResultsGrid's export view (#419): the columns, rows, column choices and
// selected records it reports to the toolbar for CSV, Copy as TSV, the
// Columns menu and Open in Dataverse.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterAll, afterEach, beforeAll, beforeEach, describe, expect, it } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { ResultsGrid, type ResultsGridProps } from './ResultsGrid';
import { toCsv, type GridExportView } from './exportResults';
import { mergePage } from './pagedResults';
import { SettingsProvider, usePersistedSettings } from '../../host/usePersistedSettings';
import type { DisplaySettings } from '../../state/displaySettings';
import { __clearLiveTables, __registerLiveTable } from '../../metadata/metadata';
import { registerFixtureTables } from '../../test/fixtures';
import type { ExecResult } from '../../engine/dataverseExecutor';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

// A ResizeObserver that reports a size, so the virtualized body renders rows.
const realResizeObserver = globalThis.ResizeObserver;
beforeAll(() => {
  globalThis.ResizeObserver = class {
    constructor(private cb: ResizeObserverCallback) {}
    observe(target: Element) {
      this.cb(
        [{ target, contentRect: { width: 1200, height: 600 } } as unknown as ResizeObserverEntry],
        this as unknown as ResizeObserver,
      );
    }
    unobserve() {}
    disconnect() {}
  } as unknown as typeof ResizeObserver;
});
afterAll(() => {
  globalThis.ResizeObserver = realResizeObserver;
});

const FV = '@OData.Community.Display.V1.FormattedValue';
const A = '00000000-0000-0000-0000-0000000000a1';
const B = '00000000-0000-0000-0000-0000000000b2';
const C = '00000000-0000-0000-0000-0000000000c3';

const account = (id: string, name: string, extra: Record<string, unknown> = {}) => ({
  accountid: id,
  name,
  statecode: 0,
  [`statecode${FV}`]: 'Active',
  ...extra,
});
const page = (...value: Record<string, unknown>[]) => ({ value });

let root: Root | null = null;
let el: HTMLDivElement | null = null;
let views: GridExportView[] = [];
let settings: [DisplaySettings, (s: DisplaySettings) => void] | null = null;

function Probe() {
  settings = usePersistedSettings();
  return null;
}

function render(props: Partial<ResultsGridProps> & { body: unknown }) {
  act(() =>
    root!.render(
      <FluentProvider theme={webLightTheme}>
        <SettingsProvider>
          <Probe />
          <ResultsGrid table="account" onViewChange={(v) => views.push(v)} {...props} />
        </SettingsProvider>
      </FluentProvider>,
    ),
  );
}
const view = () => views[views.length - 1];
const setSettings = (patch: Partial<DisplaySettings>) =>
  act(() => settings![1]({ ...settings![0], ...patch }));
/** The debounced search applies after 200 ms. */
const settle = () => act(async () => await new Promise((r) => setTimeout(r, 260)));
const tick = (name: string) => {
  const box = el!.querySelector<HTMLInputElement>(`input[aria-label="Select ${name}"]`);
  expect(box, `row checkbox for ${name}`).not.toBeNull();
  act(() => box!.click());
};
const footer = () => el!.textContent ?? '';

beforeEach(() => {
  registerFixtureTables();
  views = [];
  el = document.createElement('div');
  document.body.appendChild(el);
  root = createRoot(el);
});
afterEach(() => {
  act(() => root?.unmount());
  el?.remove();
  root = null;
  __clearLiveTables();
});

describe('ResultsGrid export view: columns (#292, #419)', () => {
  it('CSV drops hidden columns and keeps "(Raw)" twins next to their source', () => {
    render({ body: page(account(A, 'Alpha'), account(B, 'Beta')) });
    setSettings({ valueDisplayMode: 'both' });
    // Both: a formatted column that differs gets a raw twin.
    expect(view().columns.map((c) => c.id)).toEqual([
      'accountid',
      'name',
      'statecode',
      'statecode__raw',
    ]);

    setSettings({ hiddenColumns: { account: ['accountid', 'name'] } });
    expect(view().columns.map((c) => [c.id, c.sourceCol])).toEqual([
      ['statecode', 'statecode'],
      ['statecode__raw', 'statecode'],
    ]);
    const csv = toCsv(view().rows, view().columns, 'shown').slice(1).split('\r\n');
    // (No metadata for statecode in the fixture, so the header is its logical name.)
    expect(csv.slice(0, 3)).toEqual(['statecode,statecode (Raw)', 'Active,0', 'Active,0']);

    // Hiding the source hides its twin too.
    setSettings({ hiddenColumns: { account: ['statecode'] } });
    expect(view().columns.map((c) => c.id)).toEqual(['accountid', 'name']);
  });

  it('choices agree with the columns shown, including the show-everything fallback', () => {
    render({ body: page(account(A, 'Alpha')) });
    setSettings({ hiddenColumns: { account: ['name'] } });
    expect(view().choices!.map((c) => [c.key, c.shown])).toEqual([
      ['accountid', true],
      ['name', false],
      ['statecode', true],
    ]);

    // Every column this response has is hidden: the grid shows them all
    // rather than nothing, and the menu must say so.
    setSettings({ hiddenColumns: { account: ['accountid', 'name', 'statecode', 'other'] } });
    expect(view().columns.map((c) => c.id)).toEqual(['accountid', 'name', 'statecode']);
    expect(view().choices!.every((c) => c.shown)).toBe(true);
  });
});

describe('ResultsGrid export view: Open in Dataverse (#293, #419)', () => {
  it('de-duplicates a record fanned out by a collection $expand', () => {
    const contacts = [
      { contactid: C, fullname: 'Jane' },
      { contactid: B, fullname: 'John' },
    ];
    render({
      body: page(account(A, 'Alpha', { contact_customer_accounts: contacts })),
      expand: [{ nav: 'contact_customer_accounts', select: ['fullname'] }] as never,
    });
    // Two grid lines for one record; tick both.
    const boxes = [...el!.querySelectorAll<HTMLInputElement>('input[aria-label="Select Alpha"]')];
    expect(boxes).toHaveLength(2);
    act(() => boxes[0].click());
    act(() => boxes[1].click());
    expect(view().selectedRecordIds).toEqual([A]);
    expect(view().selectedRecords).toEqual([{ id: A, table: 'account' }]);
  });

  it('only counts listed rows: a search that hides the ticked row drops it (#413)', async () => {
    // The same response throughout: only the search changes.
    const body = page(account(A, 'Alpha'), account(B, 'Beta'));
    render({ body });
    tick('Alpha');
    expect(view().selectedRecordIds).toEqual([A]);
    expect(footer()).toContain('1 selected');

    render({ body, searchQuery: 'Beta' });
    await settle();
    expect(view().rows.map((r) => r.name)).toEqual(['Beta']);
    // Not Alpha, which is no longer listed.
    expect(view().selectedRecordIds).toEqual([]);
    expect(footer()).not.toContain('selected');

    // Ticking the one listed row makes it the one to open.
    tick('Beta');
    expect(view().selectedRecordIds).toEqual([B]);
    expect(footer()).toContain('1 selected');
  });

  it('is off for rows without a primary ID ($apply) and for intersect tables', () => {
    render({ body: page({ name: 'Alpha', count: 3 }, { name: 'Beta', count: 1 }) });
    expect(view().canOpenRecords).toBe(false);

    act(() => root!.render(<></>));
    __registerLiveTable({
      logicalName: 'accountleads',
      entitySetName: 'accountleadscollection',
      displayName: 'Account Leads',
      primaryKey: 'accountleadid',
      primaryName: 'name',
      columns: [],
      navigationProperties: [],
      isIntersect: true,
    });
    render({ body: page({ accountleadid: A, name: 'x' }), table: 'accountleads' });
    expect(view().canOpenRecords).toBe(false);
  });

  it('opens an activity as its own type, and only when the row says which', () => {
    __registerLiveTable({
      logicalName: 'activitypointer',
      entitySetName: 'activitypointers',
      displayName: 'Activity',
      primaryKey: 'activityid',
      primaryName: 'subject',
      columns: [],
      navigationProperties: [],
    });
    render({ body: page({ activityid: A, subject: 'Call' }), table: 'activitypointer' });
    expect(view().canOpenRecords).toBe(false);

    render({
      body: page({ activityid: B, subject: 'Mail', activitytypecode: 'email' }),
      table: 'activitypointer',
    });
    expect(view().canOpenRecords).toBe(true);
    tick('Mail');
    expect(view().selectedRecords).toEqual([{ id: B, table: 'email' }]);
  });
});

describe('ResultsGrid selection across pages (#418)', () => {
  const run = (body: unknown): ExecResult => ({
    status: 200,
    statusText: 'OK',
    ms: 1,
    bytes: 10,
    ok: true,
    outcome: 'ok',
    body,
    headers: {},
  });

  it('keeps the selection when another page is merged in, and clears it on a new run', () => {
    const first = run({ ...page(account(A, 'Alpha')), '@odata.nextLink': 'https://x/next' });
    render({ body: first.body });
    tick('Alpha');
    expect(view().selectedRecordIds).toEqual([A]);

    // Infinite scroll merges page 2: a new body, same run.
    const merged = mergePage(first, { rows: [account(B, 'Beta')], bytes: 10 });
    render({ body: merged.body });
    expect(view().rows).toHaveLength(2);
    expect(view().selectedRecordIds).toEqual([A]);
    expect(footer()).toContain('1 selected');

    // A new run with a row at the same position: not selected.
    render({ body: page(account(C, 'Gamma')) });
    expect(view().selectedRecordIds).toEqual([]);
    expect(footer()).not.toContain('selected');
  });
});

describe('ResultsGrid row search with hidden columns (#422)', () => {
  it('says hidden columns are not searched', async () => {
    render({ body: page(account(A, 'Alpha'), account(B, 'Beta')), searchQuery: 'a' });
    await settle();
    expect(footer()).toContain('Filtered');
    expect(footer()).not.toContain('hidden columns not searched');

    setSettings({ hiddenColumns: { account: ['statecode'] } });
    await settle();
    expect(footer()).toContain('hidden columns not searched');
  });
});
