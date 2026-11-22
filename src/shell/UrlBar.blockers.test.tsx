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

function bar(advisories: Advisory[], onExecute = vi.fn()) {
  const errors = advisories.filter((a) => a.severity === 'error');
  return (
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
    </FluentProvider>
  );
}

function mount(advisories: Advisory[], onExecute = vi.fn()) {
  host = document.createElement('div');
  document.body.appendChild(host);
  root = createRoot(host);
  act(() => root!.render(bar(advisories, onExecute)));
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

  it("doesn't reopen by itself after the blockers clear", () => {
    mount([blocker('a')]);
    act(() => send().click());
    expect(surface()).not.toBeNull();
    // The blockers clear (e.g. metadata arrives), then a new one appears.
    act(() => root!.render(bar([])));
    act(() => root!.render(bar([blocker('b')])));
    expect(surface()).toBeNull();
  });

  it('keeps the count for warnings only, and Execute enabled', () => {
    mount([warning]);
    expect(document.body.textContent).not.toContain('to fix');
    expect(send().getAttribute('aria-disabled')).toBeNull();
  });
});

describe('keyboard path into the drawer from a blocked Execute (#421)', () => {
  const pointed = (id: string): Advisory => ({ ...blocker(id), focusNode: `node-${id}` });
  function mountWithOpen(advisories: Advisory[]) {
    host = document.createElement('div');
    document.body.appendChild(host);
    root = createRoot(host);
    act(() =>
      root!.render(
        <FluentProvider theme={webLightTheme}>
          <UrlBar
            method="POST"
            url="/api/data/v9.2/accounts"
            executeVerb="Send"
            onExecute={vi.fn()}
            loading={false}
            advisories={advisories}
            onAdvisoryFocus={vi.fn()}
            disabledReason={advisories[0].title}
          />
        </FluentProvider>,
      ),
    );
  }
  const openButtons = () => [
    ...(surface()?.querySelectorAll<HTMLButtonElement>(
      'button[aria-label="Open the affected pane"]',
    ) ?? []),
  ];
  const key = (target: Element, k: string) =>
    act(() => {
      target.dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true }));
    });

  it('Enter on Execute moves focus to the first Open; Escape brings it back', () => {
    mountWithOpen([pointed('a'), pointed('b')]);
    const b = send();
    act(() => b.focus());
    key(b, 'Enter');
    expect(openButtons()).toHaveLength(2);
    expect(document.activeElement).toBe(openButtons()[0]);

    key(document.activeElement!, 'Escape');
    expect(surface()).toBeNull();
    expect(document.activeElement).toBe(send());
  });

  it('a click on Execute moves focus too; without an Open, to the drawer itself', () => {
    mountWithOpen([blocker('a')]);
    act(() => send().click());
    expect(surface()).not.toBeNull();
    expect(surface()!.contains(document.activeElement)).toBe(true);
  });

  it('opening from the chip leaves focus where it is', () => {
    mountWithOpen([pointed('a')]);
    const chip = [...document.body.querySelectorAll('button')].find((x) =>
      x.getAttribute('aria-label')?.startsWith('Advisories'),
    )!;
    act(() => chip.focus());
    act(() => chip.click());
    expect(surface()).not.toBeNull();
    expect(document.activeElement).toBe(chip);
  });
});
