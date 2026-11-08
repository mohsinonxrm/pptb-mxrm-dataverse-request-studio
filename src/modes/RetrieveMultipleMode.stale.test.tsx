// Retrieve Multiple works from current data, not from a render that has
// since been superseded:
//   • its advisories (and the Execute blocker) are worked out again when the
//     table's metadata arrives after a saved request was loaded;
//   • the Saved library's load runs with the current table list, even when
//     the save context was published before that list arrived.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';

const host = vi.hoisted(() => ({
  liveTable: undefined as unknown,
  entities: [] as Array<{ logicalName: string }>,
  notify: vi.fn(),
}));
vi.mock('../host/useLiveMetadata', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../host/useLiveMetadata')>()),
  useLiveTable: () => ({ table: host.liveTable, loading: false, error: null }),
  useWarmReferencedTables: () => ({ version: 0, pending: [], failed: [] }),
}));
vi.mock('../host/useScopedEntities', () => ({
  useScopedEntities: () => ({
    entities: host.entities,
    allEntities: host.entities,
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
vi.mock('../shell/dialogs', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../shell/dialogs')>()),
  notify: host.notify,
}));

import { RetrieveMultipleMode } from './RetrieveMultipleMode';
import { SaveContextRoot, type SaveContextValue } from '../state/SaveContext';
import { serializeRetrieveMultiple, type SavedRequest } from '../state/savedRequests';
import { __clearLiveTables, findTable } from '../mock/metadata';
import { group, registerFixtureTables, retrieveMultipleState } from '../test/fixtures';
import type { FilterNode } from '../editors/filter/filterTree';

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

/** A saved Retrieve Multiple request on account. */
function saved(filter: FilterNode[] = []): SavedRequest {
  return {
    id: 'saved-1',
    name: 'Saved one',
    modeId: 'retrieve-multiple',
    orgScope: 'default',
    savedAt: 1,
    state: serializeRetrieveMultiple(retrieveMultipleState({ filter: group('and', filter) })),
  } as SavedRequest;
}

const executeButton = () =>
  [...hostEl!.querySelectorAll('button')].find((b) => b.textContent?.trim() === 'Execute');

beforeEach(() => {
  __clearLiveTables();
  host.liveTable = undefined;
  host.entities = [];
  host.notify.mockReset();
  published = null;
  hostEl = document.createElement('div');
  document.body.appendChild(hostEl);
  root = createRoot(hostEl);
});
afterEach(() => {
  act(() => root?.unmount());
  hostEl?.remove();
  root = null;
});

describe('RetrieveMultipleMode', () => {
  it('works out the advisories again when the table arrives after a saved request loads', () => {
    // An Execute blocker that needs the table's metadata to be found:
    // Dataverse rejects a query function on a related table's column.
    const fnOnNavPath = {
      id: 'f1',
      type: 'function',
      op: 'EqualUserId',
      col: 'primarycontactid/ownerid',
    } as FilterNode;
    render();
    act(() => published!.onLoadSaved(saved([fnOnNavPath])));
    render();
    // No metadata yet: nothing can be checked, so nothing blocks.
    expect(executeButton()?.disabled).toBe(false);

    registerFixtureTables();
    host.liveTable = findTable('account');
    render();
    expect(executeButton()?.disabled).toBe(true);
  });

  it('loads with the current table list, not the one from when the context was published', () => {
    render();
    act(() => published!.onLoadSaved(saved()));
    render();
    const before = published!;

    // The table list arrives (another org's: no account) with nothing else changing.
    host.entities = [{ logicalName: 'contact' }];
    render();
    const loader = published!.onLoadSaved;
    expect(published).toBe(before); // not re-published: only the data is memoised on
    act(() => loader({ ...saved(), id: 'saved-2' }));
    expect(host.notify).toHaveBeenCalledWith(
      expect.stringContaining("entity `account` isn't available in this environment"),
    );
  });
});
