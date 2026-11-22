// Write results' open buttons (#293, #419, #422): Open in Dataverse after a
// write that leaves the record, never after a Delete, and the raw Web API
// URL button is named so the two aren't confused.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';

vi.mock('../../env/environment', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../../env/environment')>()),
  ENV: { name: 'Test', host: 'org.crm.dynamics.com', apiBase: '/api/data/v9.2', user: '' },
}));

import { WriteResultCard, type WriteResultContext } from './WriteResultCard';
import type { ExecResult } from '../../engine/dataverseExecutor';
import { __clearLiveTables } from '../../metadata/metadata';
import { ACCOUNT_ID, registerFixtureTables } from '../../test/fixtures';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let root: Root | null = null;
let el: HTMLDivElement | null = null;
beforeEach(() => {
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

const ok204: ExecResult = {
  status: 204,
  statusText: 'No Content',
  ms: 5,
  bytes: 0,
  ok: true,
  outcome: 'ok',
  body: null,
  headers: {},
};
const render = (ctx: WriteResultContext) =>
  act(() =>
    root!.render(
      <FluentProvider theme={webLightTheme}>
        <WriteResultCard result={ok204} ctx={ctx} />
      </FluentProvider>,
    ),
  );
const buttons = () => [...el!.querySelectorAll('button')].map((b) => b.textContent?.trim());

describe('WriteResultCard open buttons', () => {
  it('after an Update: Open Web API URL and Open in Dataverse', () => {
    render({ operation: 'update', table: 'account', recordId: ACCOUNT_ID });
    expect(buttons()).toEqual(expect.arrayContaining(['Open Web API URL', 'Open in Dataverse']));
    expect(buttons()).not.toContain('Open');
  });

  it('after a Delete: neither, but the ID can still be copied', () => {
    render({ operation: 'delete', table: 'account', recordId: ACCOUNT_ID });
    expect(buttons()).toContain('Copy ID');
    expect(buttons()).not.toContain('Open in Dataverse');
    expect(buttons()).not.toContain('Open Web API URL');
  });

  it('after clearing one column the row is still there, so both stay', () => {
    render({
      operation: 'delete',
      table: 'account',
      recordId: ACCOUNT_ID,
      clearedColumn: 'telephone1',
    });
    expect(buttons()).toEqual(expect.arrayContaining(['Open Web API URL', 'Open in Dataverse']));
  });
});
