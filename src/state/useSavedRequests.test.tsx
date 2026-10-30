// The saved-request store never writes over data it couldn't read (#220).
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { startSavedStore, useSavedRequests, type UseSavedRequestsApi } from './useSavedRequests';
import type { SavedRequest } from './savedRequests';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const KEY = 'drs:saved-requests:v1:default';
const entry = (id: string, modeId = 'update', extra: Partial<SavedRequest> = {}): SavedRequest =>
  ({ id, name: id, modeId, orgScope: 'default', state: {}, savedAt: 1, ...extra }) as SavedRequest;

const roots: Root[] = [];

/**
 * Start the store as the app does (scrub, then read storage) and render one
 * consumer. Not a fresh `import()` per test: that was slow enough to time
 * out under parallel load (#319).
 */
async function mountStore(): Promise<() => UseSavedRequestsApi> {
  act(() => startSavedStore());
  let api!: UseSavedRequestsApi;
  function Probe() {
    api = useSavedRequests();
    return null;
  }
  const root = createRoot(document.createElement('div'));
  roots.push(root);
  await act(async () => root.render(<Probe />));
  return () => api;
}

beforeEach(() => localStorage.clear());
afterEach(() => {
  for (const r of roots.splice(0)) act(() => r.unmount());
});

describe('useSavedRequests with an unreadable library', () => {
  it('refuses every write, keeps the stored text and a backup, until a new library is started', async () => {
    localStorage.setItem(KEY, '[{"id":"a","name":');
    const api = await mountStore();
    expect(api().saved).toEqual([]);
    expect(api().health).toMatchObject({ unreadable: true, backedUp: true });
    expect(api().health.unreadableText).toBe('[{"id":"a","name":');

    let r!: { ok: boolean; error?: string };
    act(() => {
      r = api().save(entry('new'));
    });
    expect(r.ok).toBe(false);
    expect(r.error).toMatch(/couldn't be read/);
    act(() => api().clearAll());
    expect(localStorage.getItem(KEY)).toBe('[{"id":"a","name":');
    const backups = Object.keys(localStorage).filter((k) =>
      k.startsWith('drs:saved-requests-corrupt:v1:default:'),
    );
    expect(backups).toHaveLength(1);

    act(() => {
      r = api().startNewLibrary();
    });
    expect(r.ok).toBe(true);
    expect(api().health.unreadable).toBe(false);
    act(() => {
      r = api().save(entry('new'));
    });
    expect(r.ok).toBe(true);
    expect(JSON.parse(localStorage.getItem(KEY)!).map((e: SavedRequest) => e.id)).toEqual(['new']);
    expect(localStorage.getItem(backups[0])).toBe('[{"id":"a","name":');
  });
});

describe('useSavedRequests with items it cannot use', () => {
  it('hides them, then writes them back unchanged after every change', async () => {
    const newer = entry('newer', 'create', { schemaVersion: 2 });
    const unknownMode = entry('future', 'graphql-query');
    localStorage.setItem(KEY, JSON.stringify([entry('ok'), newer, unknownMode]));
    const api = await mountStore();
    expect(api().saved.map((e) => e.id)).toEqual(['ok']);
    expect(api().health.hidden).toBe(2);

    act(() => {
      api().save(entry('added', 'batch'));
    });
    act(() => api().remove('ok'));
    act(() => api().clearAll());
    expect(JSON.parse(localStorage.getItem(KEY)!)).toEqual([newer, unknownMode]);
  });
});

describe('useSavedRequests error after a new library (#220 review)', () => {
  it("drops the 'couldn't be read' error once another consumer started a new library", async () => {
    localStorage.setItem(KEY, 'not json');
    act(() => startSavedStore());
    const apis: UseSavedRequestsApi[] = [];
    function Probe({ i }: { i: number }) {
      apis[i] = useSavedRequests();
      return null;
    }
    const root = createRoot(document.createElement('div'));
    roots.push(root);
    await act(async () =>
      root.render(
        <>
          <Probe i={0} />
          <Probe i={1} />
        </>,
      ),
    );
    act(() => {
      apis[0].save(entry('x'));
    });
    expect(apis[0].error).toMatch(/couldn't be read/);
    act(() => {
      apis[1].startNewLibrary();
    });
    expect(apis[0].error).toBeNull();
  });
});
