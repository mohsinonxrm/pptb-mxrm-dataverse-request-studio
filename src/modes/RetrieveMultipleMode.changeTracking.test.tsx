// Retrieve Multiple change tracking (#233): Execute sends Track changes,
// the options Dataverse rejects with it block Execute, and Get changes runs
// the newest delta link and replaces the results with the changes.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';

interface ChangeTrackingProp {
  deltaLink: string;
  onGetChanges?: () => void;
  gettingChanges?: boolean;
  isDelta?: boolean;
}
const host = vi.hoisted(() => ({
  liveTable: undefined as unknown,
  retrieve: vi.fn(),
  absoluteUrl: vi.fn(),
  props: null as null | {
    requestUrl?: string;
    result?: { body: unknown } | null;
    changeTracking?: ChangeTrackingProp;
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
    runtime: { ...orig.runtime, retrieveMultiple: host.retrieve, absoluteUrl: host.absoluteUrl },
  };
});
vi.mock('../views/CodeView', () => ({ CodeView: () => null }));
vi.mock('../views/ResultsView', () => ({
  ResultsView: (p: typeof host.props) => {
    host.props = p;
    return <div data-testid="results" />;
  },
}));

import { RetrieveMultipleMode } from './RetrieveMultipleMode';
import { SaveContextRoot, type SaveContextValue } from '../state/SaveContext';
import { serializeRetrieveMultiple, type SavedRequest } from '../state/savedRequests';
import { __clearLiveTables, __registerLiveTable, findTable } from '../metadata/metadata';
import { emptyPrefer } from '../editors/PreferEditor';
import { registerFixtureTables, retrieveMultipleState } from '../test/fixtures';
import type { RetrieveMultipleState } from '../state/readState';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const ORG = 'https://contoso.crm.dynamics.com/api/data/v9.2';
const DELTA_1 = `${ORG}/accounts?$select=name&$deltatoken=1001%2108%2f22%2f2026`;
const DELTA_2 = `${ORG}/accounts?$select=name&$deltatoken=1002%2108%2f22%2f2026`;
const DELTA_3 = `${ORG}/accounts?$select=name&$deltatoken=1003%2108%2f22%2f2026`;

const ok = (body: unknown) => ({
  status: 200,
  statusText: 'OK',
  ms: 5,
  bytes: 20,
  ok: true,
  outcome: 'ok',
  body,
  headers: {},
});

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

const button = (text: string) =>
  [...hostEl!.querySelectorAll('button')].find((b) => b.textContent?.trim() === text) as
    | HTMLButtonElement
    | undefined;
const blocked = (b: HTMLButtonElement | undefined) =>
  !!b && (b.disabled || b.getAttribute('aria-disabled') === 'true');
const flush = () => act(async () => await new Promise((r) => setTimeout(r, 0)));

function load(overrides: Partial<RetrieveMultipleState>) {
  render();
  const saved = {
    id: 'saved-1',
    name: 'Saved',
    modeId: 'retrieve-multiple',
    orgScope: 'default',
    savedAt: 1,
    state: serializeRetrieveMultiple(
      retrieveMultipleState({
        select: ['name'],
        prefer: { ...emptyPrefer(), trackChanges: true },
        ...overrides,
      }),
    ),
  } as SavedRequest;
  act(() => published!.onLoadSaved(saved));
  render();
}

beforeEach(() => {
  __clearLiveTables();
  registerFixtureTables();
  __registerLiveTable({ ...findTable('account')!, changeTrackingEnabled: true });
  host.liveTable = findTable('account');
  host.props = null;
  host.retrieve.mockReset();
  host.absoluteUrl.mockReset();
  published = null;
  hostEl = document.createElement('div');
  document.body.appendChild(hostEl);
  root = createRoot(hostEl);
});
afterEach(() => {
  act(() => root?.unmount());
  hostEl?.remove();
  root = null;
  __clearLiveTables();
});

describe('RetrieveMultipleMode change tracking (#233)', () => {
  it('$top blocks Execute while Track changes is on', () => {
    load({ top: 50 });
    expect(blocked(button('Execute'))).toBe(true);
    load({ top: null });
    expect(blocked(button('Execute'))).toBe(false);
  });

  it('a table without change tracking blocks Execute', () => {
    __registerLiveTable({ ...findTable('account')!, changeTrackingEnabled: undefined });
    host.liveTable = findTable('account');
    load({ top: null });
    expect(blocked(button('Execute'))).toBe(true);
  });

  it('Get changes runs the delta link, replaces the results, and keeps the newest delta link', async () => {
    host.retrieve.mockResolvedValue(
      ok({ value: [{ accountid: 'a1', name: 'Contoso' }], '@odata.deltaLink': DELTA_1 }),
    );
    load({ top: null, prefer: { ...emptyPrefer(), trackChanges: true, maxpagesize: 25 } });
    await act(async () => button('Execute')!.click());
    await flush();

    // Execute sent Track changes; the results offer the delta link.
    expect(host.retrieve).toHaveBeenCalledTimes(1);
    expect(host.retrieve.mock.calls[0][0].prefer.trackChanges).toBe(true);
    expect(host.props?.changeTracking).toMatchObject({ deltaLink: DELTA_1, isDelta: false });

    const deleted = {
      '@odata.context': `${ORG}/$metadata#accounts/$deletedEntity`,
      id: 'a2',
      reason: 'deleted',
    };
    host.absoluteUrl.mockResolvedValueOnce(
      ok({
        value: [{ accountid: 'a1', name: 'Contoso Ltd' }, deleted],
        '@odata.deltaLink': DELTA_2,
      }),
    );
    await act(async () => host.props!.changeTracking!.onGetChanges!());
    await flush();

    // The delta link ran with the run's page size, without track-changes.
    expect(host.absoluteUrl).toHaveBeenCalledTimes(1);
    expect(host.absoluteUrl.mock.calls[0][0]).toBe(DELTA_1);
    expect(host.absoluteUrl.mock.calls[0][1]).toBe(25);
    expect(host.absoluteUrl.mock.calls[0][3]).toBeUndefined();
    // The changes replaced the results; the next delta link replaced the old.
    expect((host.props?.result?.body as { value: unknown[] }).value).toHaveLength(2);
    expect(host.props?.requestUrl).toBe(DELTA_1);
    expect(host.props?.changeTracking).toMatchObject({ deltaLink: DELTA_2, isDelta: true });

    // And Get changes again runs the newest one.
    host.absoluteUrl.mockResolvedValueOnce(ok({ value: [], '@odata.deltaLink': DELTA_3 }));
    await act(async () => host.props!.changeTracking!.onGetChanges!());
    await flush();
    expect(host.absoluteUrl.mock.calls[1][0]).toBe(DELTA_2);
    expect(host.props?.changeTracking?.deltaLink).toBe(DELTA_3);
  });

  it('a failed Get changes keeps the delta link to try again', async () => {
    host.retrieve.mockResolvedValue(ok({ value: [], '@odata.deltaLink': DELTA_1 }));
    load({ top: null });
    await act(async () => button('Execute')!.click());
    await flush();
    host.absoluteUrl.mockResolvedValueOnce({
      ...ok({ error: { message: 'Token expired' } }),
      status: 400,
      ok: false,
      outcome: 'error',
    });
    await act(async () => host.props!.changeTracking!.onGetChanges!());
    await flush();
    expect(host.props?.result).toMatchObject({ ok: false });
    expect(host.props?.changeTracking?.deltaLink).toBe(DELTA_1);
  });

  it('a response without a delta link offers no Get changes', async () => {
    host.retrieve.mockResolvedValue(ok({ value: [] }));
    load({ top: null, prefer: emptyPrefer() });
    await act(async () => button('Execute')!.click());
    await flush();
    expect(host.props?.changeTracking).toBeUndefined();
  });
});
