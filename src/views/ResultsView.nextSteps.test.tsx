// ResultsView next steps (#291): the actions on 0 rows, and on 401 / 403 /
// 404 and every other error.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';

const host = vi.hoisted(() => ({
  copy: vi.fn(async () => true),
  entities: [
    { logicalName: 'account', entitySetName: 'accounts', displayName: 'Account', metadataId: 'a' },
    { logicalName: 'contact', entitySetName: 'contacts', displayName: 'Contact', metadataId: 'c' },
  ],
}));
vi.mock('../shell/copy', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../shell/copy')>()),
  copyText: host.copy,
}));
// A connected org, so URLs are absolute (without one they read https:///…).
vi.mock('../mock/environment', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../mock/environment')>()),
  ENV: { name: 'Test', host: 'org.crm.dynamics.com', apiBase: '/api/data/v9.2', user: '' },
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
import type { ExecResult } from '../engine/dataverseExecutor';
import { ENV } from '../mock/environment';

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
const failed = (status: number, message: string, code?: string): ExecResult => ({
  status,
  statusText: '',
  ms: 10,
  bytes: 0,
  ok: false,
  outcome: 'error',
  body: { error: { code, message } },
  headers: { Authorization: 'Bearer secret-token' },
});

describe('ResultsView next steps (#291)', () => {
  it('0 rows: shows the actions the mode wires, and each does what it says', () => {
    const onEditFilter = vi.fn();
    const onRunWithoutFilter = vi.fn();
    const ok: ExecResult = {
      status: 200,
      statusText: 'OK',
      ms: 5,
      bytes: 20,
      ok: true,
      outcome: 'ok',
      body: { value: [] },
      headers: {},
    };
    render(
      <ResultsView
        result={ok}
        mode="multi"
        table="account"
        requestUrl="/api/data/v9.2/accounts?$filter=name eq 'x'"
        onEditFilter={onEditFilter}
        onRunWithoutFilter={onRunWithoutFilter}
      />,
    );
    expect(hostEl!.textContent).toContain('No records matched');
    act(() => button('Edit $filter')!.click());
    expect(onEditFilter).toHaveBeenCalledOnce();
    act(() => button('Run without $filter')!.click());
    expect(onRunWithoutFilter).toHaveBeenCalledOnce();
    act(() => button('Copy request URL')!.click());
    expect(host.copy).toHaveBeenCalledWith(
      `https://${ENV.host}/api/data/v9.2/accounts?$filter=name eq 'x'`,
      'the request URL',
    );
  });

  it('0 rows without filter actions: only Copy request URL', () => {
    render(
      <ResultsView
        result={{
          status: 200,
          statusText: 'OK',
          ms: 5,
          bytes: 20,
          ok: true,
          outcome: 'ok',
          body: { value: [] },
          headers: {},
        }}
        mode="multi"
        requestUrl="/api/data/v9.2/accounts"
      />,
    );
    expect(button('Edit $filter')).toBeUndefined();
    expect(button('Run without $filter')).toBeUndefined();
    expect(button('Copy request URL')).toBeDefined();
  });

  it('403: names the missing privilege', () => {
    render(
      <ResultsView
        result={failed(
          403,
          "Principal user (Id=1) is missing prvReadAccount privilege (Id=2) on OTC=1 for entity 'account'.",
          '0x80042f09',
        )}
        mode="multi"
      />,
    );
    expect(hostEl!.textContent).toContain(
      'Your user is missing the prvReadAccount privilege (Read access on account).',
    );
  });

  it('401: points at reconnecting in ToolBox', () => {
    render(<ResultsView result={failed(401, 'Unauthorized')} mode="multi" />);
    expect(hostEl!.textContent).toContain('Open Connection Manager in Power Platform ToolBox');
  });

  it('404 on an entity set: suggests a scoped table, and Use hands it back', () => {
    const onUse = vi.fn();
    render(
      <ResultsView
        result={failed(0, "Resource not found for the segment 'acounts'.", '0x8006088a')}
        mode="next"
        onUseEntitySet={onUse}
      />,
    );
    expect(hostEl!.textContent).toContain('Did you mean accounts (Account)?');
    act(() => button('Use accounts')!.click());
    expect(onUse).toHaveBeenCalledWith('acounts', 'accounts');
  });

  it('404 suggestion without a handler is text only', () => {
    render(
      <ResultsView
        result={failed(404, "Resource not found for the segment 'acounts'.")}
        mode="multi"
      />,
    );
    expect(hostEl!.textContent).toContain('Did you mean accounts');
    expect(button('Use accounts')).toBeUndefined();
  });

  it('every error: Copy error report has the request, and masks a quoted token', () => {
    render(
      <ResultsView
        result={failed(400, 'Bad filter near Bearer eyJ0eXAi.secret-token', '0x80060888')}
        mode="multi"
        requestUrl="/api/data/v9.2/accounts?$filter=bad"
      />,
    );
    act(() => button('Copy error report')!.click());
    const [report, what] = host.copy.mock.calls[0] as unknown as [string, string];
    expect(what).toBe('the error report');
    expect(report).toContain(
      'Request: GET https://org.crm.dynamics.com/api/data/v9.2/accounts?$filter=bad',
    );
    expect(report).toContain('Code: 0x80060888');
    expect(report).toContain('Message: Bad filter near Bearer [removed]');
    expect(report).not.toMatch(/secret-token|eyJ0eXAi|Authorization/);
  });

  it('keeps the notice above an error, so a failed Run without $filter can be undone', () => {
    render(
      <ResultsView
        result={failed(0, 'Dataverse response exceeds the maximum size (16 MB)')}
        mode="multi"
        notice={<button type="button">Restore $filter</button>}
      />,
    );
    expect(button('Restore $filter')).toBeDefined();
    expect(hostEl!.textContent).toContain('Request failed');
  });

  it('offers no 0-row actions when the mode wires none', () => {
    render(
      <ResultsView
        result={{
          status: 200,
          statusText: 'OK',
          ms: 5,
          bytes: 20,
          ok: true,
          outcome: 'ok',
          body: { value: [] },
          headers: {},
        }}
        mode="next"
      />,
    );
    expect(hostEl!.textContent).toContain('No records matched');
    for (const t of ['Edit $filter', 'Run without $filter', 'Copy request URL']) {
      expect(button(t)).toBeUndefined();
    }
  });
});

describe('ResultsView throttling (#269)', () => {
  it('names the limit, says what to do, and offers Try again', () => {
    const onRefresh = vi.fn();
    render(
      <ResultsView
        result={failed(
          429,
          'Number of requests exceeded the limit of 6000 over time window of 300 seconds.',
          '0x80072322',
        )}
        mode="multi"
        onRefresh={onRefresh}
      />,
    );
    expect(hostEl!.textContent).toContain('Throttled: number of requests · 429');
    expect(hostEl!.textContent).toContain('about 6,000 requests per user in 5 minutes');
    expect(hostEl!.textContent).toContain('other apps and flows signed in as you count too');
    act(() => button('Try again')!.click());
    expect(onRefresh).toHaveBeenCalledOnce();
    const link = [...hostEl!.querySelectorAll('a')].find(
      (a) => a.textContent === 'About the limits',
    );
    expect(link?.getAttribute('href')).toContain('/data-platform/api-limits');
  });

  it('Try again waits for the retry that is running', () => {
    render(
      <ResultsView
        result={failed(
          429,
          'Number of concurrent requests exceeded the limit of 52.',
          '0x80072326',
        )}
        mode="multi"
        onRefresh={vi.fn()}
        loading
      />,
    );
    expect(button('Try again')).toBeUndefined();
    expect(button('Trying again…')?.disabled).toBe(true);
  });

  it('offers no Try again on other errors', () => {
    render(<ResultsView result={failed(400, 'Bad request')} mode="multi" onRefresh={vi.fn()} />);
    expect(button('Try again')).toBeUndefined();
  });
});
