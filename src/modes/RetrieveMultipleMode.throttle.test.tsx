// Retrieve all when Dataverse throttles a page (#269): it pauses with
// Continue instead of failing, keeps the rows so far, and Continue resumes
// from the refused page.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';

const host = vi.hoisted(() => ({
  liveTable: undefined as unknown,
  retrieve: vi.fn(),
  page: vi.fn(),
  props: null as null | {
    onRetrieveAll?: () => void;
    onLoadMore?: () => void;
    onContinueLoadMore?: () => void;
    loadMoreNotice?: string | null;
    loadMoreError?: string | null;
    result?: { body?: { value?: unknown[] } } | null;
  },
}));
vi.mock('../host/useLiveMetadata', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../host/useLiveMetadata')>()),
  useLiveTable: () => ({ table: host.liveTable, loading: false, error: null }),
  useWarmReferencedTables: () => ({ version: 0, pending: [], failed: [] }),
}));
vi.mock('../host/useScopedEntities', () => ({
  useScopedEntities: () => ({
    entities: [{ logicalName: 'account' }],
    allEntities: [{ logicalName: 'account' }],
    allLoading: false,
    loading: false,
    error: null,
    retry: () => {},
    needsSetup: false,
    scopeMode: 'all',
  }),
}));
vi.mock('../engine/runtime', async (importOriginal) => {
  const orig = await importOriginal<typeof import('../engine/runtime')>();
  return {
    ...orig,
    runtime: { ...orig.runtime, retrieveMultiple: host.retrieve, absoluteUrl: host.page },
  };
});
vi.mock('../views/CodeView', () => ({ CodeView: () => null }));
vi.mock('../views/ResultsView', () => ({
  ResultsView: (p: typeof host.props) => {
    host.props = p;
    return null;
  },
}));

import { RetrieveMultipleMode } from './RetrieveMultipleMode';
import { SaveContextRoot, type SaveContextValue } from '../state/SaveContext';
import { serializeRetrieveMultiple, type SavedRequest } from '../state/savedRequests';
import { __clearLiveTables, findTable } from '../mock/metadata';
import { registerFixtureTables, retrieveMultipleState } from '../test/fixtures';
import { mapError } from '../engine/dataverseExecutor';
import { clearModeDrafts } from '../state/modeDrafts';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let root: Root | null = null;
let hostEl: HTMLDivElement | null = null;
let published: SaveContextValue | null = null;

function render() {
  act(() =>
    root!.render(
      <FluentProvider theme={webLightTheme}>
        <SaveContextRoot value={published} setValue={(v) => (published = v)}>
          <RetrieveMultipleMode themeMode="light" />
        </SaveContextRoot>
      </FluentProvider>,
    ),
  );
}
const flush = () => act(async () => await new Promise((r) => setTimeout(r, 0)));
/** Run a paging handler and let its pages (and their state updates) settle. */
const page = (fn: () => void) =>
  act(async () => {
    fn();
    await new Promise((r) => setTimeout(r, 150));
  });
const okPage = (rows: number, nextLink?: string) => ({
  status: 200,
  statusText: 'OK',
  ms: 5,
  bytes: 10,
  ok: true,
  outcome: 'ok',
  body: {
    value: Array.from({ length: rows }, (_, i) => ({ accountid: `${nextLink}-${i}` })),
    ...(nextLink ? { '@odata.nextLink': nextLink } : {}),
  },
  headers: {},
});
const throttled = () =>
  mapError(
    new Error(
      'Dataverse queryData failed: 0x80072322: Number of requests exceeded the limit of 6000 over time window of 300 seconds.',
    ),
    3,
  );

beforeEach(() => {
  clearModeDrafts();
  __clearLiveTables();
  registerFixtureTables();
  host.liveTable = findTable('account');
  host.props = null;
  host.retrieve.mockReset();
  host.page.mockReset();
  host.retrieve.mockResolvedValue(okPage(2, 'https://org/next-1'));
  published = null;
  hostEl = document.createElement('div');
  document.body.appendChild(hostEl);
  root = createRoot(hostEl);
  render();
  act(() =>
    published!.onLoadSaved({
      id: 's',
      name: 'All accounts',
      modeId: 'retrieve-multiple',
      orgScope: 'default',
      savedAt: 1,
      state: serializeRetrieveMultiple(retrieveMultipleState()),
    } as SavedRequest),
  );
  render();
});
afterEach(() => {
  act(() => root?.unmount());
  hostEl?.remove();
  root = null;
  __clearLiveTables();
});

async function execute() {
  act(() =>
    [...hostEl!.querySelectorAll('button')]
      .find((b) => b.textContent?.trim() === 'Execute')!
      .click(),
  );
  await flush();
}

describe('Retrieve all, throttled (#269)', () => {
  it('pauses with Continue, keeps the rows, and resumes from the refused page', async () => {
    await execute();
    host.page
      .mockResolvedValueOnce(okPage(2, 'https://org/next-2'))
      .mockResolvedValueOnce(throttled());
    await page(() => host.props!.onRetrieveAll!());
    expect(host.props!.loadMoreError).toBeNull();
    expect(host.props!.loadMoreNotice).toMatch(
      /^Retrieve all paused: Dataverse is throttling requests \(number of requests\)\. 4 rows are loaded\./,
    );
    expect(host.props!.result?.body?.value).toHaveLength(4);

    // Continue: the refused page is asked for again.
    host.page.mockResolvedValueOnce(okPage(1));
    await page(() => host.props!.onContinueLoadMore!());
    expect(host.page.mock.calls[host.page.mock.calls.length - 1]?.[0]).toBe('https://org/next-2');
    expect(host.props!.loadMoreNotice).toBeNull();
    expect(host.props!.result?.body?.value).toHaveLength(5);
  });

  it('a throttled next page says so', async () => {
    await execute();
    host.page.mockResolvedValueOnce(throttled());
    await page(() => host.props!.onLoadMore!());
    expect(host.props!.loadMoreError).toMatch(
      /^Couldn't load the next page: Dataverse is throttling requests \(number of requests\)\./,
    );
  });
});
