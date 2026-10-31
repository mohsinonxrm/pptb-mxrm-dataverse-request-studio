// Retrieve Multiple count only (#272): advice about clauses it does not send
// stays out of the way, Execute sends countOnly, and the number reaches the view.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import type { ReactNode } from 'react';

const host = vi.hoisted(() => ({
  liveTable: undefined as unknown,
  retrieve: vi.fn(),
  props: null as null | {
    requestUrl?: string;
    onEditFilter?: () => void;
    onRunWithoutFilter?: () => void;
    notice?: ReactNode;
    result?: { body: unknown } | null;
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
  return { ...orig, runtime: { ...orig.runtime, retrieveMultiple: host.retrieve } };
});
vi.mock('../views/CodeView', () => ({ CodeView: () => null }));
// The real ResultsView is covered by its own tests; here we only need the
// props the mode hands it, and its notice rendered.
vi.mock('../views/ResultsView', () => ({
  ResultsView: (p: typeof host.props) => {
    host.props = p;
    return <div data-testid="results">{p?.notice}</div>;
  },
}));

import { RetrieveMultipleMode } from './RetrieveMultipleMode';
import { SaveContextRoot, type SaveContextValue } from '../state/SaveContext';
import { serializeRetrieveMultiple, type SavedRequest } from '../state/savedRequests';
import { __clearLiveTables, findTable } from '../mock/metadata';
import { makeExpandSpec } from '../editors/ExpandEditor';
import { emptyApply } from '../editors/ApplyEditor';
import { group, registerFixtureTables, retrieveMultipleState, rule } from '../test/fixtures';
import type { RetrieveMultipleState } from '../state/readState';

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

const button = (text: string) =>
  [...hostEl!.querySelectorAll('button')].find((b) => b.textContent?.trim() === text);
const flush = () => act(async () => await new Promise((r) => setTimeout(r, 0)));

beforeEach(() => {
  __clearLiveTables();
  registerFixtureTables();
  host.liveTable = findTable('account');
  host.props = null;
  host.retrieve.mockReset();
  host.retrieve.mockResolvedValue({
    status: 200,
    statusText: 'OK',
    ms: 5,
    bytes: 20,
    ok: true,
    outcome: 'ok',
    body: { value: [] },
    headers: {},
  });
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

/** A 1:N $expand with its own nested $expand plus $top: Dataverse rejects
 *  that $top, so the mode blocks Execute — unless only the count is sent. */
async function load(countOnly: boolean) {
  render();
  const saved = {
    id: 'saved-1',
    name: 'Saved',
    modeId: 'retrieve-multiple',
    orgScope: 'default',
    savedAt: 1,
    state: serializeRetrieveMultiple(
      retrieveMultipleState({
        countOnly,
        top: 50,
        expand: [
          {
            ...makeExpandSpec('contact_customer_accounts'),
            nestedExpand: [makeExpandSpec('parentcustomerid_account')],
          },
        ],
      }),
    ),
  } as SavedRequest;
  act(() => published!.onLoadSaved(saved));
  render();
}
const executeButton = () => button('Execute') as HTMLButtonElement | undefined;

describe('RetrieveMultipleMode count only (#272)', () => {
  it('a rows request with the forbidden $top stays blocked', async () => {
    await load(false);
    expect(executeButton()?.disabled).toBe(true);
    expect(hostEl!.textContent).not.toContain('Count only: just $filter is sent');
  });

  it('count only drops that blocker, says what is sent, and runs', async () => {
    host.retrieve.mockResolvedValue({
      status: 200,
      statusText: 'OK',
      ms: 5,
      bytes: 2,
      ok: true,
      outcome: 'ok',
      body: 42,
      headers: {},
    });
    await load(true);
    expect(executeButton()?.disabled).toBe(false);
    act(() => executeButton()!.click());
    await flush();
    expect(host.retrieve).toHaveBeenCalledOnce();
    expect((host.retrieve.mock.calls[0][0] as RetrieveMultipleState).countOnly).toBe(true);
    expect(host.props?.result?.body).toBe(42);
    expect(host.props?.requestUrl).toContain('/accounts?$select=accountid&$top=1&$count=true');
    // A count of matching rows isn't "no records": no 0-row next steps.
    expect(host.props?.onRunWithoutFilter).toBeUndefined();
  });
});

async function loadState(s: RetrieveMultipleState) {
  render();
  const saved = {
    id: 'saved-2',
    name: 'Saved',
    modeId: 'retrieve-multiple',
    orgScope: 'default',
    savedAt: 1,
    state: serializeRetrieveMultiple(s),
  } as SavedRequest;
  act(() => published!.onLoadSaved(saved));
  render();
}
const openNode = (label: string) =>
  act(() => {
    const el = [...hostEl!.querySelectorAll('aside *')].find(
      (e) => e.children.length === 0 && e.textContent?.trim() === label,
    ) as HTMLElement | undefined;
    (el?.closest('button,[role=treeitem],[role=button],li') as HTMLElement | null)?.click();
  });
const countOnlySwitch = () =>
  hostEl!.querySelector<HTMLInputElement>('input[aria-label="Count only"]')!;

describe('RetrieveMultipleMode count only and saved requests (#272)', () => {
  it('switching Count only on and off leaves an older saved request unchanged', async () => {
    // Saved before #272: no countOnly key at all.
    const { countOnly: _drop, ...older } = retrieveMultipleState({ top: 10 });
    await loadState(older as RetrieveMultipleState);
    expect(published?.dirty).toBe(false);
    openNode('$count');
    act(() => countOnlySwitch().click());
    expect(published?.dirty).toBe(true);
    act(() => countOnlySwitch().click());
    expect(published?.dirty).toBe(false);
  });
});

describe('RetrieveMultipleMode count only with $apply on (#272)', () => {
  const withApply = () =>
    retrieveMultipleState({
      countOnly: true,
      filter: group('and', [rule('name', 'eq', 'Counted')]),
      apply: {
        ...emptyApply(),
        enabled: true,
        prefilter: group('and', [rule('name', 'eq', 'Aggregated')]),
      },
    });

  it('sends $filter, not the $apply pipeline', async () => {
    await loadState(withApply());
    act(() => button('Execute')!.click());
    await flush();
    expect(host.props?.requestUrl).toContain("$filter=name eq 'Counted'");
    expect(host.props?.requestUrl).not.toContain('$apply');
    expect(host.props?.requestUrl).not.toContain('Aggregated');
  });

  it("doesn't call the $filter ignored", async () => {
    await loadState(withApply());
    openNode('$filter');
    expect(hostEl!.textContent).toContain('Combine rules with AND / OR');
    expect(hostEl!.textContent).not.toContain('is ignored while $apply is active');
  });
});
