// #298: the header shows the saved request's name, "Unsaved" while it has
// changes, and renames it in place.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { SavedNameTag } from './SavedNameTag';
import { reloadSavedFromStorage } from '../state/useSavedRequests';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const KEY = 'drs:saved-requests:v1:default';
const entry = {
  id: 'acc',
  name: 'Active accounts',
  modeId: 'retrieve-multiple',
  orgScope: 'default',
  state: {},
  savedAt: 1,
};
let root: Root | null = null;
let host: HTMLDivElement | null = null;

function render(props: { lastSavedId?: string; dirty: boolean }) {
  localStorage.setItem(KEY, JSON.stringify([entry]));
  act(() => reloadSavedFromStorage());
  host = document.createElement('div');
  document.body.appendChild(host);
  root = createRoot(host);
  act(() =>
    root!.render(
      <FluentProvider theme={webLightTheme}>
        <SavedNameTag {...props} />
      </FluentProvider>,
    ),
  );
  return host;
}

beforeEach(() => localStorage.clear());
afterEach(() => {
  act(() => root?.unmount());
  host?.remove();
});

describe('SavedNameTag (#298)', () => {
  it('shows nothing for a request that was never saved', () => {
    expect(render({ dirty: true }).textContent).toBe('');
  });

  it('shows the name, and Unsaved only while there are changes', () => {
    expect(render({ lastSavedId: 'acc', dirty: false }).textContent).toBe('Active accounts');
    act(() => root!.unmount());
    expect(render({ lastSavedId: 'acc', dirty: true }).textContent).toBe('Active accountsUnsaved');
  });

  it('renames on Enter', () => {
    const h = render({ lastSavedId: 'acc', dirty: false });
    act(() => h.querySelector('button')!.click());
    const input = h.querySelector('input')!;
    const setValue = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!;
    act(() => {
      setValue.call(input, 'Open accounts');
      input.dispatchEvent(new Event('input', { bubbles: true }));
    });
    act(() => {
      input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    });
    expect(h.textContent).toBe('Open accounts');
    expect(JSON.parse(localStorage.getItem(KEY)!)[0].name).toBe('Open accounts');
  });
});
