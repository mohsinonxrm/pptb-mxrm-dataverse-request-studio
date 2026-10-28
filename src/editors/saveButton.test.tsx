// #308: Save is available again once the entry a request came from is deleted.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const KEY = 'drs:saved-requests:v1:default';
let root: Root | null = null;
let host: HTMLDivElement | null = null;

async function renderSave(props: { dirty: boolean; lastSavedId?: string }) {
  vi.resetModules(); // the store reads storage at import
  const { SaveButton } = await import('./SavedRequestsDialog');
  host = document.createElement('div');
  document.body.appendChild(host);
  root = createRoot(host);
  await act(async () =>
    root!.render(
      <SaveButton state={{}} modeId="retrieve-multiple" onSaved={() => {}} {...props} />,
    ),
  );
  // The tooltip relabels the button, so find it by its text.
  return [...host.querySelectorAll('button')].find(
    (b) => b.textContent?.trim() === 'Save',
  ) as HTMLButtonElement;
}

// The first dynamic import of the dialog module is slow in jsdom.
vi.setConfig({ testTimeout: 30_000 });

beforeEach(() => localStorage.clear());
afterEach(() => {
  act(() => root?.unmount());
  host?.remove();
  root = null;
  host = null;
});

const entry = {
  id: 'kept',
  name: 'kept',
  modeId: 'retrieve-multiple',
  orgScope: 'default',
  state: {},
  savedAt: 1,
};

describe('SaveButton', () => {
  it('is disabled for a clean request whose saved entry still exists', async () => {
    localStorage.setItem(KEY, JSON.stringify([entry]));
    expect((await renderSave({ dirty: false, lastSavedId: 'kept' })).disabled).toBe(true);
  });

  it('is enabled for a clean request whose saved entry was deleted', async () => {
    localStorage.setItem(KEY, JSON.stringify([]));
    expect((await renderSave({ dirty: false, lastSavedId: 'gone' })).disabled).toBe(false);
  });

  it('stays disabled for a clean request that was never saved', async () => {
    expect((await renderSave({ dirty: false })).disabled).toBe(true);
  });

  it("stays disabled when the library can't be read", async () => {
    localStorage.setItem(KEY, '[{"id":"a","name":');
    expect((await renderSave({ dirty: false, lastSavedId: 'gone' })).disabled).toBe(true);
  });
});
