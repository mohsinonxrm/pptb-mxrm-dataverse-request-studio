// The results toolbar over a real grid (#419): Open in Dataverse for one
// selected row (#293), Copy as TSV following the grid (#415), and the
// summary's total when Dataverse stopped counting (#414).
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterAll, afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';

const host = vi.hoisted(() => ({
  copy: vi.fn(async (_text: string, _what: string, _opts?: unknown) => true),
  open: vi.fn(async (_url: string) => {}),
}));
vi.mock('../shell/copy', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../shell/copy')>()),
  copyText: host.copy,
}));
vi.mock('../host/openInDataverse', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../host/openInDataverse')>()),
  openInBrowser: host.open,
}));
vi.mock('../env/environment', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../env/environment')>()),
  ENV: { name: 'Test', host: 'org.crm.dynamics.com', apiBase: '/api/data/v9.2', user: '' },
}));

import { ResultsView } from './ResultsView';
import { SettingsProvider, usePersistedSettings } from '../host/usePersistedSettings';
import type { DisplaySettings } from '../state/displaySettings';
import type { ExecResult } from '../engine/dataverseExecutor';
import { __clearLiveTables } from '../metadata/metadata';
import { registerFixtureTables } from '../test/fixtures';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

// A ResizeObserver that reports a size, so the virtualized body renders rows.
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
const B = '00000000-0000-0000-0000-0000000000b2';
const FV = '@OData.Community.Display.V1.FormattedValue';

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
const accounts = (extra: Record<string, unknown> = {}) => ({
  value: [
    { accountid: A, name: 'Alpha', statecode: 0, [`statecode${FV}`]: 'Active' },
    { accountid: B, name: 'Beta', statecode: 1, [`statecode${FV}`]: 'Inactive' },
  ],
  ...extra,
});

let root: Root | null = null;
let el: HTMLDivElement | null = null;
let settings: [DisplaySettings, (s: DisplaySettings) => void] | null = null;
function Probe() {
  settings = usePersistedSettings();
  return null;
}

function render(result: ExecResult, table = 'account') {
  act(() =>
    root!.render(
      <FluentProvider theme={webLightTheme}>
        <SettingsProvider>
          <Probe />
          <ResultsView result={result} table={table} mode="multi" requestUrl="/accounts" />
        </SettingsProvider>
      </FluentProvider>,
    ),
  );
}
const button = (text: string) =>
  [...document.body.querySelectorAll('button')].find((b) => b.textContent?.trim() === text);
const tick = (name: string) =>
  act(() => el!.querySelector<HTMLInputElement>(`input[aria-label="Select ${name}"]`)!.click());

beforeEach(() => {
  registerFixtureTables();
  host.copy.mockClear();
  host.open.mockClear();
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

describe('Open in Dataverse in the results toolbar (#293, #419)', () => {
  it('opens the one selected row', () => {
    render(ok(accounts()));
    const open = button('Open in Dataverse')!;
    expect(open.disabled || open.getAttribute('aria-disabled') === 'true').toBe(true);

    tick('Beta');
    expect(button('Open in Dataverse')!.disabled).toBe(false);
    act(() => button('Open in Dataverse')!.click());
    expect(host.open).toHaveBeenCalledWith(
      `https://org.crm.dynamics.com/main.aspx?pagetype=entityrecord&etn=account&id=${B}`,
    );

    // Two rows: which one to open isn't clear.
    tick('Alpha');
    expect(button('Open in Dataverse')!.disabled).toBe(true);
  });

  it("isn't offered for rows without a primary ID ($apply)", () => {
    render(ok({ value: [{ name: 'Alpha', total: 3 }] }));
    expect(button('Open in Dataverse')).toBeUndefined();
  });
});

describe('Copy as TSV (#415)', () => {
  it('copies the shown columns and listed rows, as the grid shows them', async () => {
    render(ok(accounts()));
    act(() => settings![1]({ ...settings![0], hiddenColumns: { account: ['accountid'] } }));
    act(() => button('Copy data')!.click());
    const item = [...document.body.querySelectorAll('[role=menuitem]')].find((m) =>
      m.textContent?.includes('Copy as TSV'),
    ) as HTMLElement;
    act(() => item.click());
    expect(host.copy).toHaveBeenCalledTimes(1);
    const [text, what] = host.copy.mock.calls[0];
    expect(text.split('\n')).toEqual(['Name\tstatecode', 'Alpha\tActive', 'Beta\tInactive']);
    expect(what).toBe('2 rows');
  });
});

describe('Summary total (#414)', () => {
  const summary = () =>
    el!.querySelector('[data-testid=results-summary]')!.textContent!.replace(/\s+/g, ' ');

  it('reads "5,000+" when Dataverse stopped counting', () => {
    render(
      ok(
        accounts({
          '@odata.count': 5000,
          '@Microsoft.Dynamics.CRM.totalrecordcountlimitexceeded': true,
          '@odata.nextLink': 'https://org.crm.dynamics.com/api/data/v9.2/accounts?$skiptoken=x',
        }),
      ),
    );
    expect(summary()).toContain(`total ${(5000).toLocaleString()}+`);
  });

  it('reads the plain total when it is exact', () => {
    render(
      ok(
        accounts({
          '@odata.count': 2,
          '@Microsoft.Dynamics.CRM.totalrecordcountlimitexceeded': false,
        }),
      ),
    );
    expect(summary()).toContain('total 2');
    expect(summary()).not.toContain('+');
  });
});
