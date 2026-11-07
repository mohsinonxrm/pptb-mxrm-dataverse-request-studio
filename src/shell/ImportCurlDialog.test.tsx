// Import cURL dialog (#271): a read that finishes after the text changed or
// the dialog closed is dropped, so Open never opens a command that isn't shown.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import type { CurlImportResult } from '../engine/curlImport/curlImport';

const reads = vi.hoisted(() => ({
  pending: [] as Array<(r: CurlImportResult) => void>,
}));
vi.mock('../engine/curlImport/curlImport', () => ({
  importCurl: () => new Promise<CurlImportResult>((resolve) => reads.pending.push(resolve)),
}));
vi.mock('../host/useScopedEntities', () => ({
  useScopedEntities: () => ({ allEntities: [] }),
}));

import { ImportCurlDialog } from './ImportCurlDialog';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let root: Root | null = null;
let hostEl: HTMLDivElement | null = null;

const result = (path: string): CurlImportResult => ({
  request: { modeId: 'batch', groups: [], continueOnError: false },
  method: 'GET',
  path,
  notes: [],
  warnings: [],
  errors: [],
});

function render(open = true) {
  act(() =>
    root!.render(
      <FluentProvider theme={webLightTheme}>
        <ImportCurlDialog open={open} onClose={() => render(false)} />
      </FluentProvider>,
    ),
  );
}

const textarea = () =>
  document.querySelector<HTMLTextAreaElement>('textarea[aria-label="cURL command"]')!;
const button = (text: string) =>
  [...document.querySelectorAll('button')].find((b) => b.textContent?.trim().startsWith(text));
function type(value: string) {
  act(() => {
    Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!.call(
      textarea(),
      value,
    );
    textarea().dispatchEvent(new Event('input', { bubbles: true }));
  });
}

beforeEach(() => {
  reads.pending = [];
  hostEl = document.createElement('div');
  document.body.appendChild(hostEl);
  root = createRoot(hostEl);
});
afterEach(() => {
  act(() => root?.unmount());
  hostEl?.remove();
});

describe('ImportCurlDialog', () => {
  it('shows the result of the text that was read', async () => {
    render();
    type('curl a');
    act(() => button('Read command')!.click());
    await act(async () => reads.pending[0](result('a-path')));
    expect(document.body.textContent).toContain('a-path');
    expect(button('Open in')).toBeTruthy();
  });

  it('drops a read that finishes after the text changed', async () => {
    render();
    type('curl a');
    act(() => button('Read command')!.click());
    type('curl b');
    await act(async () => reads.pending[0](result('a-path')));
    expect(document.body.textContent).not.toContain('a-path');
    expect(button('Open in')).toBeUndefined();
    expect(button('Read command')!.hasAttribute('disabled')).toBe(false);
  });

  it('drops a read that finishes after Cancel', async () => {
    render();
    type('curl a');
    act(() => button('Read command')!.click());
    act(() => button('Cancel')!.click());
    await act(async () => reads.pending[0](result('a-path')));
    render(true);
    expect(textarea().value).toBe('');
    expect(document.body.textContent).not.toContain('a-path');
  });
});
