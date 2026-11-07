// #295: with blockers, the chip reads "N to fix" and the blocked Execute opens
// the advisories instead of doing nothing.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { UrlBar } from './UrlBar';
import type { Advisory } from '../primitives/advisories';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let root: Root | null = null;
let host: HTMLDivElement | null = null;
afterEach(() => {
  act(() => root?.unmount());
  host?.remove();
  root = null;
  host = null;
});

const blocker = (id: string): Advisory => ({
  id,
  severity: 'error',
  source: 'validation',
  title: `Blocker ${id}`,
});
const warning: Advisory = { id: 'w', severity: 'warning', source: 'validation', title: 'Careful' };

function mount(advisories: Advisory[], onExecute = vi.fn()) {
  host = document.createElement('div');
  document.body.appendChild(host);
  root = createRoot(host);
  const errors = advisories.filter((a) => a.severity === 'error');
  act(() =>
    root!.render(
      <FluentProvider theme={webLightTheme}>
        <UrlBar
          method="POST"
          url="/api/data/v9.2/accounts"
          executeVerb="Send"
          onExecute={onExecute}
          loading={false}
          advisories={advisories}
          disabledReason={errors.length ? errors[0].title : null}
        />
      </FluentProvider>,
    ),
  );
  return onExecute;
}

const send = () =>
  [...document.body.querySelectorAll('button')].find(
    (b) => b.textContent?.trim() === 'Send',
  ) as HTMLButtonElement;
const surface = () => document.body.querySelector('[role=dialog], .fui-PopoverSurface');

describe('blocked Execute (#295)', () => {
  it('reads "N to fix" and opens the drawer from a click', () => {
    const onExecute = mount([blocker('a'), blocker('b'), warning]);
    expect(document.body.textContent).toContain('2 to fix');
    const b = send();
    expect(b.getAttribute('aria-disabled')).toBe('true');
    expect(b.disabled).toBe(false); // still focusable
    expect(surface()).toBeNull();
    act(() => b.click());
    expect(surface()?.textContent).toContain('2 things to fix before you can send');
    expect(onExecute).not.toHaveBeenCalled();
  });

  it('opens from Enter on the focused button', () => {
    mount([blocker('a')]);
    const b = send();
    act(() => {
      b.focus();
      b.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    });
    expect(surface()?.textContent).toContain('1 thing to fix before you can send');
  });

  it('keeps the count for warnings only, and Execute enabled', () => {
    mount([warning]);
    expect(document.body.textContent).not.toContain('to fix');
    expect(send().getAttribute('aria-disabled')).toBeNull();
  });
});
