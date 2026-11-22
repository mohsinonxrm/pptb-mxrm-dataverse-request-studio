// #225 PR 2: PUT of one column (Update), clearing one column (Delete) and
// Upsert by alternate key run now. The advisories that blocked Execute are
// gone, Execute is enabled, and the pane where the shape is picked says the
// request goes through ToolBox's $batch.
import { act, type ComponentType } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import type { Advisory } from '../primitives/advisories';

const host = vi.hoisted(() => ({
  liveTable: undefined as unknown,
  urlBar: null as null | { advisories?: Advisory[]; disabledReason?: string | null },
  sidebar: null as null | { onSelect: (id: string) => void },
}));
vi.mock('../shell/UrlBar', () => ({
  UrlBar: (p: typeof host.urlBar) => {
    host.urlBar = p;
    return null;
  },
}));
vi.mock('../shell/Sidebar', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../shell/Sidebar')>()),
  Sidebar: (p: typeof host.sidebar) => {
    host.sidebar = p;
    return null;
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
vi.mock('../views/CodeView', () => ({ CodeView: () => null }));
vi.mock('../views/ResultsView', () => ({ ResultsView: () => null }));

import { UpdateMode } from './UpdateMode';
import { DeleteMode } from './DeleteMode';
import { UpsertMode } from './UpsertMode';
import { SaveContextRoot, type SaveContextValue } from '../state/SaveContext';
import {
  getOrgScope,
  serializeDelete,
  serializeUpdate,
  serializeUpsert,
} from '../state/savedRequests';
import { clearModeDrafts, writeModeDraft } from '../state/modeDrafts';
import { BATCH_TRANSPORT_NOTE } from '../primitives/BatchTransportNote';
import { __clearLiveTables, findTable } from '../metadata/metadata';
import { emptyPrefer } from '../editors/PreferEditor';
import { defaultWriteHeaders } from '../editors/HeadersEditor';
import { ACCOUNT_ID, registerFixtureTables } from '../test/fixtures';
import {
  defaultBypassOptions,
  type DeleteState,
  type UpdateState,
  type UpsertState,
} from '../state/writeState';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
globalThis.ResizeObserver ??= class {
  observe() {}
  unobserve() {}
  disconnect() {}
} as unknown as typeof ResizeObserver;

let root: Root | null = null;
let hostEl: HTMLDivElement | null = null;
let published: SaveContextValue | null = null;

beforeEach(() => {
  __clearLiveTables();
  registerFixtureTables();
  host.liveTable = findTable('account');
  host.urlBar = null;
  host.sidebar = null;
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
  clearModeDrafts();
});

function render(Mode: ComponentType) {
  act(() =>
    root!.render(
      <FluentProvider theme={webLightTheme}>
        <SaveContextRoot value={published} setValue={(v) => (published = v)}>
          <Mode />
        </SaveContextRoot>
      </FluentProvider>,
    ),
  );
}

/** Start the mode from `state` (as a draft kept from an earlier visit) and open the pane `node`. */
function open(Mode: ComponentType, modeId: string, state: unknown, node: string) {
  writeModeDraft(getOrgScope(), modeId, { state, lastSavedId: undefined, lastSavedHash: null });
  render(Mode);
  act(() => host.sidebar!.onSelect(node));
}

const text = () => hostEl!.textContent ?? '';
const errorIds = () =>
  (host.urlBar?.advisories ?? []).filter((a) => a.severity === 'error').map((a) => a.id);

// One component per mode: an inline one would remount (and reset) on each render.
const Update = () => <UpdateMode themeMode="light" />;
const Delete = () => <DeleteMode themeMode="light" />;
const Upsert = () => <UpsertMode themeMode="light" />;

describe('the three $batch shapes can be executed (#225)', () => {
  it('Update, PUT one column: no blocker, Execute enabled, the note on the Method pane', () => {
    const s: UpdateState = {
      table: 'account',
      recordId: ACCOUNT_ID,
      method: 'PUT',
      putColumn: 'name',
      fieldValues: { name: 'Contoso' },
      nullFields: [],
      rawBody: null,
      prefer: emptyPrefer(),
      headers: defaultWriteHeaders(),
      returnSelect: [],
      concurrency: { kind: 'update-only' },
      bypass: defaultBypassOptions(),
      dirty: new Set(),
    };
    open(Update, 'update', serializeUpdate(s), 'method');
    expect(errorIds()).not.toContain('upd-put-unsupported');
    expect(host.urlBar?.disabledReason ?? null).toBeNull();
    expect(text()).toContain(BATCH_TRANSPORT_NOTE);
    expect(text()).not.toMatch(/isn't supported by the PPTB host/);
  });

  it('Delete, clear one column: no blocker, Execute enabled, the note on the Scope pane', () => {
    const s: DeleteState = {
      table: 'account',
      recordId: ACCOUNT_ID,
      scope: { kind: 'single-property', column: 'accountnumber' },
      concurrency: { kind: 'none' },
      headers: defaultWriteHeaders(),
      bypassCustomPlugins: false,
      bypass: defaultBypassOptions(),
      confirmText: '',
      acknowledged: false,
      dirty: new Set(),
    };
    open(Delete, 'delete', serializeDelete(s), 'scope');
    expect(errorIds()).not.toContain('del-single-property-unsupported');
    expect(host.urlBar?.disabledReason ?? null).toBeNull();
    expect(text()).toContain(BATCH_TRANSPORT_NOTE);
    expect(text()).not.toContain('Author-only inside PPTB');
    // Once, not repeated.
    expect(text().split(BATCH_TRANSPORT_NOTE)).toHaveLength(2);
  });

  it('Upsert by alternate key: no blocker, Execute enabled, the note on the Key pane', () => {
    const s: UpsertState = {
      table: 'account',
      key: {
        kind: 'alternate',
        keyName: 'akey_accountnumber',
        keyValues: { accountnumber: 'A-1' },
      },
      fieldValues: { name: 'Fabrikam' },
      nullFields: [],
      rawBody: null,
      prefer: emptyPrefer(),
      headers: defaultWriteHeaders(),
      returnSelect: [],
      concurrency: { kind: 'none' },
      bypass: defaultBypassOptions(),
      dirty: new Set(),
    };
    open(Upsert, 'upsert', serializeUpsert(s), 'key');
    expect(errorIds()).not.toContain('ups-altkey-unsupported');
    expect(host.urlBar?.disabledReason ?? null).toBeNull();
    expect(text()).toContain(BATCH_TRANSPORT_NOTE);
  });

  it('Upsert by GUID: no $batch note (it goes through the typed update)', () => {
    const s: UpsertState = {
      table: 'account',
      key: { kind: 'guid', recordId: ACCOUNT_ID },
      fieldValues: { name: 'Fabrikam' },
      nullFields: [],
      rawBody: null,
      prefer: emptyPrefer(),
      headers: defaultWriteHeaders(),
      returnSelect: [],
      concurrency: { kind: 'none' },
      bypass: defaultBypassOptions(),
      dirty: new Set(),
    };
    open(Upsert, 'upsert', serializeUpsert(s), 'key');
    expect(text()).not.toContain(BATCH_TRANSPORT_NOTE);
  });
});

describe("one column: the Precondition pane says why there's none (#225)", () => {
  // Live: Dataverse refuses any If-Match on a column's URL (405, 0x80060888).
  it("Update, PUT: the reason instead of the choices, and an ETag problem doesn't block", () => {
    const s: UpdateState = {
      table: 'account',
      recordId: ACCOUNT_ID,
      method: 'PUT',
      putColumn: 'name',
      fieldValues: { name: 'Contoso' },
      nullFields: [],
      rawBody: null,
      prefer: emptyPrefer(),
      headers: defaultWriteHeaders(),
      returnSelect: [],
      concurrency: { kind: 'etag', etag: 'not an etag' },
      bypass: defaultBypassOptions(),
      dirty: new Set(),
    };
    open(Update, 'update', serializeUpdate(s), 'precondition');
    expect(text()).toContain('Not sent with a single-column PUT');
    expect(hostEl!.querySelector('input[type=radio]')).toBeNull();
    expect(host.urlBar?.disabledReason ?? null).toBeNull();
  });

  it('Update, PATCH: the choices as before', () => {
    const s: UpdateState = {
      table: 'account',
      recordId: ACCOUNT_ID,
      method: 'PATCH',
      putColumn: null,
      fieldValues: { name: 'Contoso' },
      nullFields: [],
      rawBody: null,
      prefer: emptyPrefer(),
      headers: defaultWriteHeaders(),
      returnSelect: [],
      concurrency: { kind: 'update-only' },
      bypass: defaultBypassOptions(),
      dirty: new Set(),
    };
    open(Update, 'update', serializeUpdate(s), 'precondition');
    expect(text()).not.toContain('Not sent with a single-column PUT');
    expect(hostEl!.querySelectorAll('input[type=radio]').length).toBe(3);
  });

  it('Delete, clear one column: the reason instead of the choices', () => {
    const s: DeleteState = {
      table: 'account',
      recordId: ACCOUNT_ID,
      scope: { kind: 'single-property', column: 'accountnumber' },
      concurrency: { kind: 'etag', etag: 'W/"1"' },
      headers: defaultWriteHeaders(),
      bypassCustomPlugins: false,
      bypass: defaultBypassOptions(),
      confirmText: '',
      acknowledged: false,
      dirty: new Set(),
    };
    open(Delete, 'delete', serializeDelete(s), 'precondition');
    expect(text()).toContain('Not sent when clearing one column');
    expect(hostEl!.querySelector('input[type=radio]')).toBeNull();
  });
});
