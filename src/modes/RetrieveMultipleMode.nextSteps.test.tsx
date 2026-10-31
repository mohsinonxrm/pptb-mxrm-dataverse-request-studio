// Retrieve Multiple's next steps on 0 rows (#291): Edit $filter opens the
// $filter node, Run without $filter clears it from the builder and runs,
// and Restore puts it back.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import type { ReactNode } from 'react';

const host = vi.hoisted(() => ({
  liveTable: undefined as unknown,
  retrieve: vi.fn(),
  props: null as null | {
    onEditFilter?: () => void;
    onRunWithoutFilter?: () => void;
    notice?: ReactNode;
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
const sentFilterRules = (call: number) =>
  (host.retrieve.mock.calls[call][0] as RetrieveMultipleState).filter.rules.length;
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

async function runWithFilter() {
  render();
  const saved = {
    id: 'saved-1',
    name: 'Saved',
    modeId: 'retrieve-multiple',
    orgScope: 'default',
    savedAt: 1,
    state: serializeRetrieveMultiple(
      retrieveMultipleState({ filter: group('and', [rule('name', 'eq', 'Nobody')]) }),
    ),
  } as SavedRequest;
  act(() => published!.onLoadSaved(saved));
  render();
  act(() => button('Execute')!.click());
  await flush();
}

describe('RetrieveMultipleMode next steps (#291)', () => {
  it('offers Edit $filter and Run without $filter when a $filter is set', async () => {
    await runWithFilter();
    expect(host.retrieve).toHaveBeenCalledTimes(1);
    expect(sentFilterRules(0)).toBe(1);
    expect(host.props?.onEditFilter).toBeTypeOf('function');
    expect(host.props?.onRunWithoutFilter).toBeTypeOf('function');
    expect(host.props?.notice).toBeFalsy();
  });

  it('Edit $filter goes back to the builder', async () => {
    await runWithFilter();
    act(() => host.props!.onEditFilter!());
    expect(hostEl!.querySelector('[data-testid=results]')).toBeNull();
  });

  it('Run without $filter clears it, runs, and Restore puts it back', async () => {
    await runWithFilter();
    act(() => host.props!.onRunWithoutFilter!());
    await flush();
    expect(host.retrieve).toHaveBeenCalledTimes(2);
    expect(sentFilterRules(1)).toBe(0);
    // No filter left to edit or drop; the notice offers Restore instead.
    expect(host.props?.onRunWithoutFilter).toBeUndefined();
    expect(hostEl!.textContent).toContain('Ran without $filter');

    act(() => button('Restore $filter')!.click());
    expect(hostEl!.textContent).not.toContain('Ran without $filter');
    expect(host.props?.onRunWithoutFilter).toBeTypeOf('function');

    // The next run sends the restored $filter.
    act(() => button('Execute')!.click());
    await flush();
    expect(sentFilterRules(2)).toBe(1);
  });

  it('a normal run drops the Restore offer', async () => {
    await runWithFilter();
    act(() => host.props!.onRunWithoutFilter!());
    await flush();
    expect(hostEl!.textContent).toContain('Ran without $filter');
    act(() => button('Execute')!.click());
    await flush();
    expect(hostEl!.textContent).not.toContain('Ran without $filter');
  });
});
