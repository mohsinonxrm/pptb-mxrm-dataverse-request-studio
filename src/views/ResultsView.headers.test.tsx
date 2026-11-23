// The Headers tab (#225 PR 3): a request sent through $batch shows its
// part's real headers, the batch's own x-ms-* headers in their own group, and
// offers the returned ETag as the next Send's If-Match. Typed results keep the
// "not exposed" text.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { ResultsView, responseEtag } from './ResultsView';
import { SettingsProvider } from '../host/usePersistedSettings';
import type { ExecResult } from '../engine/dataverseExecutor';
import { __clearLiveTables } from '../metadata/metadata';
import { registerFixtureTables } from '../test/fixtures';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
globalThis.ResizeObserver ??= class {
  observe() {}
  unobserve() {}
  disconnect() {}
} as unknown as typeof ResizeObserver;

const ETAG = 'W/"18778200"';
const NOT_EXPOSED = "Response headers aren't exposed by the PPTB host";

const result = (over: Partial<ExecResult>): ExecResult => ({
  status: 200,
  statusText: 'OK',
  ms: 12,
  bytes: 128,
  ok: true,
  outcome: 'ok',
  body: { accountid: '00000000-0000-0000-0000-0000000000a1', name: 'Fabrikam' },
  headers: {},
  ...over,
});

let root: Root | null = null;
let el: HTMLDivElement | null = null;
beforeEach(() => {
  __clearLiveTables();
  registerFixtureTables();
  el = document.createElement('div');
  document.body.appendChild(el);
  root = createRoot(el);
});
afterEach(() => {
  act(() => root?.unmount());
  el?.remove();
  __clearLiveTables();
});

function render(r: ExecResult, onUseEtag?: (etag: string) => void) {
  act(() =>
    root!.render(
      <FluentProvider theme={webLightTheme}>
        <SettingsProvider>
          <ResultsView result={r} mode="single" table="account" onUseEtag={onUseEtag} />
        </SettingsProvider>
      </FluentProvider>,
    ),
  );
  const tab = [...el!.querySelectorAll<HTMLElement>('[role=tab]')].find((t) =>
    t.textContent?.trim().startsWith('Headers'),
  )!;
  act(() => tab.click());
}

const text = () => el!.textContent ?? '';
const button = (name: string) =>
  [...el!.querySelectorAll<HTMLButtonElement>('button')].find(
    (b) => b.textContent?.trim() === name,
  );

describe('Headers tab', () => {
  it('a typed result: the "not exposed" text, no ETag action', () => {
    const use = vi.fn();
    render(result({}), use);
    expect(text()).toContain(NOT_EXPOSED);
    expect(text()).not.toContain('Sent through $batch');
    expect(button('Use as If-Match')).toBeUndefined();
  });

  it("a $batch result: the part's headers, said where they come from", () => {
    render(
      result({
        transport: 'batch',
        status: 201,
        statusText: 'Created',
        headers: { etag: ETAG, 'preference-applied': 'return=representation' },
      }),
    );
    expect(text()).toContain('HTTP/1.1 201 Created');
    expect(text()).toContain('Sent through $batch');
    expect(text()).toContain('preference-applied:');
    expect(text()).toContain('return=representation');
    expect(text()).not.toContain(NOT_EXPOSED);
  });

  it('a $batch result with no headers says Dataverse sent none', () => {
    render(result({ transport: 'batch', status: 204, statusText: 'No Content', body: null }));
    expect(text()).toContain('Dataverse sent no headers for this request.');
    expect(text()).not.toContain(NOT_EXPOSED);
  });

  it("the batch's x-ms-* headers in their own group", () => {
    render(
      result({
        transport: 'batch',
        headers: { 'odata-version': '4.0' },
        batchHeaders: { 'x-ms-ratelimit-burst-remaining-xrm-requests': '7999' },
      }),
    );
    const t = text();
    expect(t).toContain('From the $batch response');
    expect(t.indexOf('odata-version')).toBeLessThan(t.indexOf('From the $batch response'));
    expect(t.indexOf('From the $batch response')).toBeLessThan(
      t.indexOf('x-ms-ratelimit-burst-remaining-xrm-requests'),
    );
  });

  it('Use as If-Match hands the etag header to the mode, then says it is set', () => {
    const use = vi.fn();
    render(
      result({
        transport: 'batch',
        headers: { etag: ETAG },
        body: { '@odata.etag': 'W/"1"' },
      }),
      use,
    );
    const b = button('Use as If-Match')!;
    expect(b.disabled).toBe(false);
    act(() => b.click());
    expect(use).toHaveBeenCalledWith(ETAG);
    expect(text()).toContain('Precondition set');
    expect(button('Use as If-Match')!.disabled).toBe(true);
  });

  it("falls back to the row's @odata.etag when there's no etag header", () => {
    const use = vi.fn();
    render(result({ body: { '@odata.etag': ETAG, name: 'Fabrikam' } }), use);
    expect(text()).toContain('@odata.etag');
    act(() => button('Use as If-Match')!.click());
    expect(use).toHaveBeenCalledWith(ETAG);
  });

  it("no action when the mode doesn't take an ETag", () => {
    render(result({ transport: 'batch', headers: { etag: ETAG } }));
    expect(button('Use as If-Match')).toBeUndefined();
  });
});

describe('responseEtag', () => {
  it('prefers the header, then the body, else nothing', () => {
    expect(responseEtag({ etag: ` ${ETAG} ` }, { '@odata.etag': 'W/"1"' })).toEqual({
      etag: ETAG,
      from: 'header',
    });
    expect(responseEtag({}, { '@odata.etag': ETAG })).toEqual({ etag: ETAG, from: 'body' });
    expect(responseEtag({}, [{ '@odata.etag': ETAG }])).toBeNull();
    expect(responseEtag({}, null)).toBeNull();
    expect(responseEtag({ etag: '' }, { '@odata.etag': '' })).toBeNull();
  });
});
