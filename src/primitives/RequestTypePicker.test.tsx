// #289 — the picker shows each type's description as its accessible
// description (not its name), and palette search matches description text.
import { act } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, describe, expect, it } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { RequestTypePicker } from './RequestTypePicker';
import { REQ_TYPES } from '../registry/requestTypes';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
// Fluent's positioning observes size; jsdom has no ResizeObserver.
globalThis.ResizeObserver ??= class {
  observe() {}
  unobserve() {}
  disconnect() {}
} as unknown as typeof ResizeObserver;

let host: HTMLDivElement;
let root: ReturnType<typeof createRoot> | null = null;
async function mount() {
  host = document.createElement('div');
  document.body.appendChild(host);
  const r = createRoot(host);
  root = r;
  await act(async () =>
    r.render(
      <FluentProvider theme={webLightTheme}>
        <RequestTypePicker activeId="retrieve-multiple" onPick={() => {}} />
      </FluentProvider>,
    ),
  );
}
const describedBy = (el: Element) =>
  document.getElementById(el.getAttribute('aria-describedby') ?? '')?.textContent;

afterEach(async () => {
  // Unmount, or this picker's Ctrl+K listener opens a second palette next test.
  await act(async () => root?.unmount());
  root = null;
  document.body.innerHTML = '';
});

describe('RequestTypePicker descriptions', () => {
  it('names each menu item by type and verb, and describes it with the description', async () => {
    await mount();
    await act(async () => host.querySelector<HTMLButtonElement>('button[aria-haspopup]')!.click());
    const items = [...document.querySelectorAll('[role="menuitem"]')];
    expect(items).toHaveLength(REQ_TYPES.length);
    for (const r of REQ_TYPES) {
      // Menu order is by group, so find each type's item by its name.
      const item = items.find((el) => el.getAttribute('aria-label') === `${r.name}, ${r.method}`);
      expect(item, r.id).toBeDefined();
      expect(describedBy(item!), r.id).toBe(r.description);
    }
  });

  it('palette search matches description text', async () => {
    await mount();
    await act(async () => {
      window.dispatchEvent(new KeyboardEvent('keydown', { key: 'k', ctrlKey: true }));
    });
    const input = document.querySelector<HTMLInputElement>(
      'input[aria-label="Search request types"]',
    )!;
    await act(async () => {
      const set = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!;
      set.call(input, 'alternate key');
      input.dispatchEvent(new Event('input', { bubbles: true }));
    });
    const rows = [...document.querySelectorAll('[data-palette-row]')];
    expect(rows.map((r) => r.getAttribute('aria-label'))).toEqual(['Upsert Record, PATCH']);
    expect(describedBy(rows[0])).toMatch(/alternate key/);
  });
});
