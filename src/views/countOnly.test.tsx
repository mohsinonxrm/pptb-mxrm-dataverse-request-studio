// Count only (#272): the result view and the Count only switch.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';

const host = vi.hoisted(() => ({
  copy: vi.fn(async () => true),
  entities: [
    { logicalName: 'account', entitySetName: 'accounts', displayName: 'Account', metadataId: 'a' },
  ],
}));
vi.mock('../shell/copy', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../shell/copy')>()),
  copyText: host.copy,
}));
vi.mock('../host/useScopedEntities', () => ({
  useScopedEntities: () => ({
    entities: host.entities,
    allEntities: host.entities,
    allLoading: false,
    loading: false,
    error: null,
    retry: () => {},
    needsSetup: false,
    scopeMode: 'all',
  }),
}));

import { ResultsView } from './ResultsView';
import { CountEditor } from '../editors/TopCountEditors';
import type { ExecResult } from '../engine/dataverseExecutor';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
// The grid and MessageBar measure themselves; jsdom has no ResizeObserver.
globalThis.ResizeObserver ??= class {
  observe() {}
  unobserve() {}
  disconnect() {}
} as unknown as typeof ResizeObserver;

let root: Root | null = null;
let hostEl: HTMLDivElement | null = null;

beforeEach(() => {
  host.copy.mockClear();
  hostEl = document.createElement('div');
  document.body.appendChild(hostEl);
  root = createRoot(hostEl);
});
afterEach(() => {
  act(() => root?.unmount());
  hostEl?.remove();
  root = null;
});

function render(ui: React.ReactElement) {
  act(() => root!.render(<FluentProvider theme={webLightTheme}>{ui}</FluentProvider>));
}
const button = (text: string) =>
  [...hostEl!.querySelectorAll('button')].find((b) => b.textContent?.trim() === text);
const text = () => hostEl!.textContent ?? '';

const counted = (n: number, countLimitExceeded?: boolean): ExecResult => ({
  status: 200,
  statusText: 'OK',
  ms: 7,
  bytes: 2,
  ok: true,
  outcome: 'ok',
  body: n,
  headers: {},
  ...(countLimitExceeded !== undefined ? { countLimitExceeded } : {}),
});

describe('ResultsView count result (#272)', () => {
  it('shows the number, not the rows grid, and copies it', () => {
    render(<ResultsView result={counted(1234)} mode="multi" table="account" />);
    expect(text()).toContain((1234).toLocaleString());
    expect(text()).toContain('matching rows');
    expect(text()).not.toContain('No records matched');
    expect(text()).not.toContain('RetrieveTotalRecordCount');
    act(() => button('Copy count')!.click());
    expect(host.copy).toHaveBeenCalledWith('1234', 'the count');
  });

  it('says one row in the singular, and zero is a count, not an empty result', () => {
    render(<ResultsView result={counted(1)} mode="multi" table="account" />);
    expect(text()).toContain('matching row');
    expect(text()).not.toContain('matching rows');
    render(<ResultsView result={counted(0)} mode="multi" table="account" />);
    expect(text()).toContain('matching rows');
    expect(text()).not.toContain('No records matched');
  });

  it('says more rows match when Dataverse reports the limit', () => {
    render(<ResultsView result={counted(5000, true)} mode="multi" table="account" />);
    expect(text()).toMatch(/stops counting at 5,000 rows/);
    expect(text()).toContain('and more rows match');
  });

  it('trusts Dataverse when it says the limit was not hit', () => {
    render(<ResultsView result={counted(5000, false)} mode="multi" table="account" />);
    expect(text()).not.toMatch(/stops counting/);
  });

  it.each([5000, 500])('without the flag, at %i says there may be more', (n) => {
    render(<ResultsView result={counted(n)} mode="multi" table="account" />);
    expect(text()).toMatch(/stops counting at 5,000 rows/);
    expect(text()).toContain('RetrieveTotalRecordCount');
  });
});

describe('ResultsView count without @odata.count (#272)', () => {
  it('shows the error, not a count', () => {
    render(
      <ResultsView
        result={{
          status: 200,
          statusText: 'OK',
          ms: 7,
          bytes: 0,
          ok: false,
          outcome: 'error',
          body: {
            error: {
              code: 'DRS.CountUnreadable',
              message: "Dataverse answered without @odata.count, so there's no count to show.",
            },
          },
          headers: {},
        }}
        mode="multi"
        table="account"
      />,
    );
    expect(text()).toContain('without @odata.count');
    expect(text()).not.toContain('matching row');
  });
});

describe('CountEditor Count only switch (#272)', () => {
  const countOnlySwitch = () =>
    hostEl!.querySelector<HTMLInputElement>('input[aria-label="Count only"]');
  const countSwitch = () =>
    [...hostEl!.querySelectorAll<HTMLInputElement>('input[role="switch"]')].find(
      (i) => i.getAttribute('aria-label') !== 'Count only',
    );

  it('is absent unless the mode wires it', () => {
    render(<CountEditor countOn={false} setCountOn={() => {}} />);
    expect(countOnlySwitch()).toBeNull();
    expect(countSwitch()!.disabled).toBe(false);
  });

  it('turns on, and disables the $count switch while on', () => {
    const setCountOnly = vi.fn();
    render(
      <CountEditor
        countOn={false}
        setCountOn={() => {}}
        countOnly={false}
        setCountOnly={setCountOnly}
      />,
    );
    act(() => countOnlySwitch()!.click());
    expect(setCountOnly).toHaveBeenCalledWith(true);

    render(<CountEditor countOn setCountOn={() => {}} countOnly setCountOnly={setCountOnly} />);
    expect(countOnlySwitch()!.checked).toBe(true);
    expect(countSwitch()!.disabled).toBe(true);
    // Count only does send $count=true, so the switch shows it on.
    expect(countSwitch()!.checked).toBe(true);
    expect(text()).toContain('Count only includes @odata.count');
  });

  it('shows the $apply banner only for a rows request', () => {
    render(<CountEditor countOn setCountOn={() => {}} applyActive setCountOnly={() => {}} />);
    expect(text()).toContain('$count is ignored while $apply is active');
    render(
      <CountEditor countOn setCountOn={() => {}} applyActive countOnly setCountOnly={() => {}} />,
    );
    expect(text()).not.toContain('is ignored while $apply is active');
  });
});
