// Change tracking in the results (#233): the delta link with Get changes and
// Copy delta link, deleted rows in the summary, and a delta with no changes.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterAll, afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';

const host = vi.hoisted(() => ({
  copy: vi.fn(async (_text: string, _what: string, _opts?: unknown) => true),
}));
vi.mock('../shell/copy', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../shell/copy')>()),
  copyText: host.copy,
}));

import { ResultsView } from './ResultsView';
import { SettingsProvider } from '../host/usePersistedSettings';
import type { ExecResult } from '../engine/dataverseExecutor';
import { __clearLiveTables } from '../metadata/metadata';
import { registerFixtureTables } from '../test/fixtures';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const realResizeObserver = globalThis.ResizeObserver;
beforeAll(() => {
  globalThis.ResizeObserver = class {
    constructor(private cb: ResizeObserverCallback) {}
    observe(target: Element) {
      this.cb(
        [{ target, contentRect: { width: 1200, height: 600 } } as unknown as ResizeObserverEntry],
        this as unknown as ResizeObserver,
      );
    }
    unobserve() {}
    disconnect() {}
  } as unknown as typeof ResizeObserver;
});
afterAll(() => {
  globalThis.ResizeObserver = realResizeObserver;
});

const A = '00000000-0000-0000-0000-0000000000a1';
const D = '00000000-0000-0000-0000-0000000000d4';
const DELTA = 'https://org.crm.dynamics.com/api/data/v9.2/accounts?$select=name&$deltatoken=7';

const ok = (body: unknown): ExecResult => ({
  status: 200,
  statusText: 'OK',
  ms: 12,
  bytes: 2048,
  ok: true,
  outcome: 'ok',
  body,
  headers: {},
});
const deletedEntry = {
  '@odata.context': 'https://org.crm.dynamics.com/api/data/v9.2/$metadata#accounts/$deletedEntity',
  id: D,
  reason: 'deleted',
};

let root: Root | null = null;
let el: HTMLDivElement | null = null;

function render(
  result: ExecResult,
  changeTracking?: Parameters<typeof ResultsView>[0]['changeTracking'],
) {
  act(() =>
    root!.render(
      <FluentProvider theme={webLightTheme}>
        <SettingsProvider>
          <ResultsView
            result={result}
            mode="multi"
            table="account"
            select={['name']}
            changeTracking={changeTracking}
          />
        </SettingsProvider>
      </FluentProvider>,
    ),
  );
}
const button = (text: string) =>
  [...el!.querySelectorAll('button')].find((b) => b.textContent?.trim() === text) as
    | HTMLButtonElement
    | undefined;
const bar = () => el!.querySelector('[data-testid="change-tracking-bar"]');

beforeEach(() => {
  registerFixtureTables();
  host.copy.mockClear();
  el = document.createElement('div');
  document.body.appendChild(el);
  root = createRoot(el);
});
afterEach(() => {
  act(() => root?.unmount());
  el?.remove();
  root = null;
  __clearLiveTables();
});

describe('ResultsView change tracking (#233)', () => {
  it('offers Get changes and Copy delta link, and shows the delta link', () => {
    const onGetChanges = vi.fn();
    render(ok({ value: [{ accountid: A, name: 'Alpha' }], '@odata.deltaLink': DELTA }), {
      deltaLink: DELTA,
      onGetChanges,
    });
    expect(bar()?.textContent).toContain('Change tracking is on');
    expect(bar()?.textContent).toContain(DELTA);
    act(() => button('Get changes')!.click());
    expect(onGetChanges).toHaveBeenCalledTimes(1);
    act(() => button('Copy delta link')!.click());
    expect(host.copy).toHaveBeenCalledWith(DELTA, 'the delta link');
  });

  it('shows deleted rows in the grid and counts them in the summary', () => {
    render(
      ok({ value: [{ accountid: A, name: 'Alpha v2' }, deletedEntry], '@odata.deltaLink': DELTA }),
      { deltaLink: DELTA, onGetChanges: () => {}, isDelta: true },
    );
    expect(bar()?.textContent).toContain('Changes since the previous delta link');
    const summary = el!.querySelector('[data-testid="results-summary"]')!.textContent;
    expect(summary).toContain('2 rows');
    expect(summary).toContain('1 deleted');
    expect(el!.textContent).toContain('Deleted');
    expect(el!.textContent).toContain(D);
  });

  it('a delta with no changes says so', () => {
    render(ok({ value: [], '@odata.deltaLink': DELTA }), {
      deltaLink: DELTA,
      onGetChanges: () => {},
      isDelta: true,
    });
    expect(el!.textContent).toContain('No changes');
    expect(el!.textContent).toContain(
      'Nothing was created, updated or deleted since the previous delta link.',
    );
    expect(button('Get changes')).toBeDefined();
  });

  it('keeps Get changes over a failed one, and disables it while it runs', () => {
    const failed: ExecResult = {
      ...ok({ error: { message: 'Expired' } }),
      status: 400,
      ok: false,
      outcome: 'error',
    };
    render(failed, { deltaLink: DELTA, onGetChanges: () => {}, isDelta: true });
    expect(button('Get changes')?.disabled).toBe(false);
    render(failed, { deltaLink: DELTA, onGetChanges: () => {}, gettingChanges: true });
    expect(button('Getting changes…')?.disabled).toBe(true);
  });

  it('shows nothing without a delta link', () => {
    render(ok({ value: [{ accountid: A, name: 'Alpha' }] }));
    expect(bar()).toBeNull();
  });
});
