// #308: Save is available again once the entry a request came from is deleted.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { SaveButton } from './SavedRequestsDialog';
import { reloadSavedFromStorage } from '../state/useSavedRequests';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const KEY = 'drs:saved-requests:v1:default';
let root: Root | null = null;
let host: HTMLDivElement | null = null;

/**
 * Seed the library and render the button. The store is reloaded from storage
 * instead of re-importing the module per test, which was slow enough to time
 * out under parallel load (#319).
 */
function renderSave(props: { dirty: boolean; lastSavedId?: string }, stored?: string) {
  localStorage.clear();
  if (stored !== undefined) localStorage.setItem(KEY, stored);
  act(() => reloadSavedFromStorage());
  host = document.createElement('div');
  document.body.appendChild(host);
  root = createRoot(host);
  act(() =>
    root!.render(
      <SaveButton state={{}} modeId="retrieve-multiple" onSaved={() => {}} {...props} />,
    ),
  );
  // The tooltip relabels the button, so find it by its text.
  return [...host.querySelectorAll('button')].find(
    (b) => b.textContent?.trim() === 'Save',
  ) as HTMLButtonElement;
}

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
  it('is disabled for a clean request whose saved entry still exists', () => {
    const button = renderSave({ dirty: false, lastSavedId: 'kept' }, JSON.stringify([entry]));
    expect(button.disabled).toBe(true);
  });

  it('is enabled for a clean request whose saved entry was deleted', () => {
    const button = renderSave({ dirty: false, lastSavedId: 'gone' }, JSON.stringify([]));
    expect(button.disabled).toBe(false);
  });

  it('stays disabled for a clean request that was never saved', () => {
    expect(renderSave({ dirty: false }).disabled).toBe(true);
  });

  it("stays disabled when the library can't be read", () => {
    const button = renderSave({ dirty: false, lastSavedId: 'gone' }, '[{"id":"a","name":');
    expect(button.disabled).toBe(true);
  });
});
