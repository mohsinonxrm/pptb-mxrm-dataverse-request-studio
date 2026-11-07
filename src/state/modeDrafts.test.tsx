// Mode drafts: switching modes keeps an unsaved request and its Save
// baseline (#101); switching connection drops every draft, so nothing
// prepared against one environment comes back against another (#93).
import { act, type ReactNode } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
  clearModeDrafts,
  readModeDraft,
  useModeDraft,
  useStashModeDraft,
  writeModeDraft,
  type ModeDraft,
} from './modeDrafts';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

// The org scope comes from the host connection; each test sets it here.
const env = vi.hoisted(() => ({ org: 'default' }));
vi.mock('./savedRequests', async (importOriginal) => ({
  ...(await importOriginal<typeof import('./savedRequests')>()),
  getOrgScope: () => env.org,
}));

const A = 'contoso.crm.dynamics.com';
const B = 'fabrikam.crm.dynamics.com';
const draft = (state: unknown, lastSavedId?: string, lastSavedHash: string | null = null) => ({
  state,
  lastSavedId,
  lastSavedHash,
});

let root: Root | null = null;
function mount(node: ReactNode): void {
  root ??= createRoot(document.createElement('div'));
  act(() => root!.render(node));
}
function unmount(): void {
  act(() => root?.unmount());
  root = null;
}

beforeEach(() => {
  env.org = A;
  clearModeDrafts();
});
afterEach(() => {
  if (root) unmount();
});

describe('readModeDraft / writeModeDraft', () => {
  it('reads back a draft written for the active org', () => {
    writeModeDraft(A, 'create', draft({ table: 'account' }, 's1', 'h1'));
    expect(readModeDraft('create')).toEqual(draft({ table: 'account' }, 's1', 'h1'));
  });

  it('keeps one draft per mode', () => {
    writeModeDraft(A, 'create', draft('c'));
    writeModeDraft(A, 'update', draft('u'));
    writeModeDraft(A, 'create', draft('c2'));
    expect(readModeDraft('create')?.state).toBe('c2');
    expect(readModeDraft('update')?.state).toBe('u');
    expect(readModeDraft('delete')).toBeUndefined();
  });

  it('ignores a draft made against an org that is no longer active (#93)', () => {
    writeModeDraft(B, 'create', draft('from B'));
    expect(readModeDraft('create')).toBeUndefined();
  });

  it('drops every draft on a connection switch, including on switching back (#93)', () => {
    writeModeDraft(A, 'create', draft('a'));
    writeModeDraft(A, 'update', draft('a'));
    env.org = B;
    expect(readModeDraft('create')).toBeUndefined();
    writeModeDraft(B, 'create', draft('b'));
    expect(readModeDraft('create')?.state).toBe('b');
    env.org = A;
    expect(readModeDraft('create')).toBeUndefined();
    expect(readModeDraft('update')).toBeUndefined();
  });

  it('a write after a switch also drops the old org drafts', () => {
    writeModeDraft(A, 'update', draft('a'));
    env.org = B;
    writeModeDraft(B, 'create', draft('b'));
    expect(readModeDraft('update')).toBeUndefined();
    expect(readModeDraft('create')?.state).toBe('b');
  });

  it('clearModeDrafts empties the store', () => {
    writeModeDraft(A, 'create', draft('a'));
    clearModeDrafts();
    expect(readModeDraft('create')).toBeUndefined();
  });
});

describe('useModeDraft', () => {
  function Probe(props: {
    modeKey: string;
    deserialize: (s: unknown) => unknown;
    init: () => unknown;
    out: { current?: ModeDraft<unknown> };
  }) {
    props.out.current = useModeDraft(props.modeKey, props.deserialize, props.init);
    return null;
  }

  it('starts from init() with no Save baseline when there is no draft', () => {
    const out: { current?: ModeDraft<unknown> } = {};
    const deserialize = vi.fn();
    mount(<Probe modeKey="create" deserialize={deserialize} init={() => 'fresh'} out={out} />);
    expect(out.current).toEqual({ state: 'fresh', lastSavedId: undefined, lastSavedHash: null });
    expect(deserialize).not.toHaveBeenCalled();
  });

  it('restores the draft through the deserializer, with its Save baseline (#101)', () => {
    writeModeDraft(A, 'create', draft({ table: 'account' }, 'saved-1', 'hash-1'));
    const out: { current?: ModeDraft<unknown> } = {};
    const init = vi.fn(() => 'fresh');
    mount(<Probe modeKey="create" deserialize={(s) => ({ restored: s })} init={init} out={out} />);
    expect(out.current).toEqual({
      state: { restored: { table: 'account' } },
      lastSavedId: 'saved-1',
      lastSavedHash: 'hash-1',
    });
    expect(init).not.toHaveBeenCalled();
  });

  it('starts fresh when the draft cannot be rebuilt', () => {
    writeModeDraft(A, 'create', draft('broken', 'saved-1', 'hash-1'));
    const out: { current?: ModeDraft<unknown> } = {};
    const deserialize = () => {
      throw new Error('bad snapshot');
    };
    mount(<Probe modeKey="create" deserialize={deserialize} init={() => 'fresh'} out={out} />);
    expect(out.current).toEqual({ state: 'fresh', lastSavedId: undefined, lastSavedHash: null });
  });

  it('reads once: a later draft does not replace the mounted state', () => {
    const out: { current?: ModeDraft<unknown> } = {};
    const props = { modeKey: 'create', deserialize: (s: unknown) => s, init: () => 'fresh', out };
    mount(<Probe {...props} />);
    writeModeDraft(A, 'create', draft('later'));
    mount(<Probe {...props} />);
    expect(out.current?.state).toBe('fresh');
  });

  it("does not restore another org's draft after a connection switch (#93)", () => {
    writeModeDraft(A, 'create', draft('from A'));
    env.org = B;
    const out: { current?: ModeDraft<unknown> } = {};
    mount(<Probe modeKey="create" deserialize={(s) => s} init={() => 'fresh'} out={out} />);
    expect(out.current?.state).toBe('fresh');
  });
});

describe('useStashModeDraft', () => {
  function Stash(props: {
    modeKey: string;
    serialized: unknown;
    lastSavedId?: string;
    lastSavedHash?: string | null;
  }) {
    useStashModeDraft(
      props.modeKey,
      props.serialized,
      props.lastSavedId,
      props.lastSavedHash ?? null,
    );
    return null;
  }

  it('stashes nothing while the mode is mounted', () => {
    mount(<Stash modeKey="create" serialized="v1" />);
    expect(readModeDraft('create')).toBeUndefined();
  });

  it('stashes the latest request and Save baseline on unmount (#101)', () => {
    mount(<Stash modeKey="create" serialized="v1" />);
    mount(<Stash modeKey="create" serialized="v2" lastSavedId="s1" lastSavedHash="h1" />);
    unmount();
    expect(readModeDraft('create')).toEqual(draft('v2', 's1', 'h1'));
  });

  it('round-trips with useModeDraft across a mode switch', () => {
    function Mode(props: { typed?: string }) {
      const d = useModeDraft<string, string>(
        'update',
        (s) => s,
        () => '',
      );
      useStashModeDraft('update', props.typed ?? d.state, d.lastSavedId, d.lastSavedHash);
      seen = d.state;
      return null;
    }
    let seen = '';
    mount(<Mode />);
    mount(<Mode typed="half-typed request" />);
    unmount(); // switch to another mode
    mount(<Mode />); // and back
    expect(seen).toBe('half-typed request');
  });

  it('stashes under the old key when the key changes', () => {
    mount(<Stash modeKey="action:WinOpportunity" serialized="win" />);
    mount(<Stash modeKey="action:LoseOpportunity" serialized="lose" />);
    expect(readModeDraft('action:WinOpportunity')?.state).toBe('win');
    unmount();
    expect(readModeDraft('action:LoseOpportunity')?.state).toBe('lose');
  });

  it('drops a draft made before a connection switch (#93)', () => {
    mount(<Stash modeKey="create" serialized="against A" />);
    env.org = B;
    mount(<Stash modeKey="create" serialized="re-rendered after switch" />);
    unmount();
    expect(readModeDraft('create')).toBeUndefined();
    env.org = A;
    expect(readModeDraft('create')).toBeUndefined();
  });

  it('tags the draft with the org that connects after the mode mounted', () => {
    env.org = 'default';
    mount(<Stash modeKey="create" serialized="v1" />);
    env.org = A;
    mount(<Stash modeKey="create" serialized="v2" />);
    unmount();
    expect(readModeDraft('create')?.state).toBe('v2');
  });

  it('keeps a draft made with no connection while still not connected', () => {
    env.org = 'default';
    mount(<Stash modeKey="create" serialized="offline" />);
    unmount();
    expect(readModeDraft('create')?.state).toBe('offline');
  });
});
