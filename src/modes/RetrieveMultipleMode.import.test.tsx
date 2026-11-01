// cURL import (#271): an imported request reaches Retrieve Multiple whether
// the mode is already open or opens because of it, and isn't a saved entry.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';

const host = vi.hoisted(() => ({ liveTable: undefined as unknown }));
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
vi.mock('../views/CodeView', () => ({ CodeView: () => null }));
vi.mock('../views/ResultsView', () => ({ ResultsView: () => null }));

import { RetrieveMultipleMode } from './RetrieveMultipleMode';
import { SaveContextRoot, type SaveContextValue } from '../state/SaveContext';
import { openImport, useRegisterModeNavigator } from '../state/modeHandoff';
import { importCurl } from '../engine/curlImport/curlImport';
import { __clearLiveTables, findTable } from '../mock/metadata';
import { account, registerFixtureTables } from '../test/fixtures';
import type { SerializedRetrieveMultipleState } from '../state/savedRequests';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let root: Root | null = null;
let hostEl: HTMLDivElement | null = null;
let published: SaveContextValue | null = null;
const navigate = vi.fn();

function Navigator() {
  useRegisterModeNavigator(navigate);
  return null;
}

function render(withMode = true) {
  act(() =>
    root!.render(
      <FluentProvider theme={webLightTheme}>
        <SaveContextRoot value={published} setValue={(v) => (published = v)}>
          <Navigator />
          {withMode && <RetrieveMultipleMode themeMode="light" />}
        </SaveContextRoot>
      </FluentProvider>,
    ),
  );
}

async function imported() {
  const r = await importCurl(
    `curl 'https://contoso.crm.dynamics.com/api/data/v9.2/accounts?$select=name&$filter=numberofemployees gt 10&$top=5' -H 'MSCRMCallerID: 00000000-0000-0000-0000-0000000000c1'`,
    {
      host: 'contoso.crm.dynamics.com',
      resolveEntitySet: (s) => (s === 'accounts' ? 'account' : undefined),
      loadTable: async () => account,
    },
  );
  expect(r.errors).toEqual([]);
  return r.request!;
}

const current = () => published?.state as SerializedRetrieveMultipleState;

beforeEach(() => {
  __clearLiveTables();
  registerFixtureTables();
  host.liveTable = findTable('account');
  navigate.mockReset();
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

describe('RetrieveMultipleMode cURL import (#271)', () => {
  it('fills the open mode and switches to it', async () => {
    render();
    const request = await imported();
    act(() => {
      expect(openImport(request)).toBe(true);
    });
    render();
    expect(navigate).toHaveBeenCalledWith('retrieve-multiple');
    expect(current()).toMatchObject({ table: 'account', select: ['name'], top: 5 });
    expect(current().filter.rules).toHaveLength(1);
    const caller = current().headers.find((h) => h.name === 'MSCRMCallerID');
    expect(caller).toMatchObject({ value: '00000000-0000-0000-0000-0000000000c1', enabled: true });
    // Not a saved entry: Save asks for a name.
    expect(published?.lastSavedId).toBeUndefined();
    expect(published?.dirty).toBe(true);
  });

  it('waits for the mode when it is not open yet', async () => {
    render(false);
    const request = await imported();
    act(() => {
      openImport(request);
    });
    render(true);
    expect(current()).toMatchObject({ table: 'account', select: ['name'], top: 5 });
  });
});
