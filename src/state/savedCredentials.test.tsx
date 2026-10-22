// Credential headers never reach storage through any saved-request path (#217).
import { act } from 'react';
import { createRoot } from 'react-dom/client';
import { beforeEach, describe, expect, it, vi } from 'vitest';
import type { UseSavedRequestsApi } from './useSavedRequests';
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

/** Load the store fresh (it scrubs and reads storage at import) and render one consumer. */
async function mountStore(): Promise<() => UseSavedRequestsApi> {
  vi.resetModules();
  const { useSavedRequests } = await import('./useSavedRequests');
  let api!: UseSavedRequestsApi;
  function Probe() {
    api = useSavedRequests();
    return null;
  }
  const root = createRoot(document.createElement('div'));
  await act(async () => root.render(<Probe />));
  return () => api;
}

beforeEach(() => localStorage.clear());

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
