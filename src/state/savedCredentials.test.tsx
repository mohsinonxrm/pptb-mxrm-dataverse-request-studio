// Credential headers never reach storage through any saved-request path (#217).
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { startSavedStore, useSavedRequests, type UseSavedRequestsApi } from './useSavedRequests';
import type { SavedRequest } from './savedRequests';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const KEY = 'drs:saved-requests:v1:default';
const OTHER = 'drs:saved-requests:v1:other.example';
const TOKEN = 'Bearer drs-test-token';
const row = (name: string, value: string) => ({ id: name, name, value, enabled: true });
const entry = (id: string, extra: Partial<SavedRequest> = {}): SavedRequest =>
  ({
    id,
    name: id,
    modeId: 'retrieve-multiple',
    orgScope: 'default',
    state: { table: 'account', headers: [row('Authorization', TOKEN), row('MSCRMCallerID', 'c')] },
    savedAt: 1,
    ...extra,
  }) as SavedRequest;

const roots: Root[] = [];

/**
 * Start the store as the app does at load (scrub credentials, then read
 * storage) and render one consumer. Not a fresh `import()` per test: that
 * was slow enough to time out under parallel load (#319).
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

describe('saved requests never keep credential headers', () => {
  it('scrubs credentials an earlier build stored, in every connection, keeping other items as they are', async () => {
    const odd = { id: 'odd', note: 'not a saved request' };
    localStorage.setItem(KEY, JSON.stringify([entry('a'), odd]));
    localStorage.setItem(OTHER, JSON.stringify([entry('b', { orgScope: 'other.example' })]));
    localStorage.setItem('drs:saved-requests:v1:broken', '{not json');
    const api = await mountStore();
    for (const k of [KEY, OTHER]) expect(localStorage.getItem(k)).not.toContain(TOKEN);
    expect(JSON.parse(localStorage.getItem(KEY)!)[1]).toEqual(odd);
    expect(localStorage.getItem(KEY)).toContain('MSCRMCallerID');
    expect(localStorage.getItem('drs:saved-requests:v1:broken')).toBe('{not json');
    expect(JSON.stringify(api().saved)).not.toContain(TOKEN);
  });

  // The one test that loads the module afresh: it checks that loading it
  // starts the store (the scrub and the read), which the tests above call
  // by hand. A fresh import is slow under load, hence the timeout (#319).
  it('scrubs and reads storage when the module loads', async () => {
    localStorage.setItem(KEY, JSON.stringify([entry('a')]));
    vi.resetModules();
    const fresh = await import('./useSavedRequests');
    expect(localStorage.getItem(KEY)).not.toContain(TOKEN);
    let api!: UseSavedRequestsApi;
    function Probe() {
      api = fresh.useSavedRequests();
      return null;
    }
    const root = createRoot(document.createElement('div'));
    roots.push(root);
    await act(async () => root.render(<Probe />));
    expect(api.saved.map((e) => e.id)).toEqual(['a']);
  }, 30_000);

  it('strips them on save and on import', async () => {
    const api = await mountStore();
    act(() => {
      api().save(entry('saved'));
    });
    act(() => {
      api().importEntries([entry('imported')]);
    });
    const stored = localStorage.getItem(KEY)!;
    expect(stored).toContain('"saved"');
    expect(stored).toContain('"imported"');
    expect(stored).not.toContain(TOKEN);
    expect(stored).toContain('MSCRMCallerID');
  });
});
