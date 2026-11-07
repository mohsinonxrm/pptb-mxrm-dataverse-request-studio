// Update / Upsert through ToolBox (#224): the host passes no row back, so a
// 204 shows no stand-in body, and asking for return=representation says why.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
// The Retrieve this record button needs a mode picker to switch with.
vi.mock('../../state/modeHandoff', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../../state/modeHandoff')>()),
  canOpenInMode: () => true,
}));

import { WriteResultCard, type WriteResultContext } from './WriteResultCard';
import type { ExecResult } from '../../engine/dataverseExecutor';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let root: Root | null = null;
let host: HTMLDivElement | null = null;
beforeEach(() => {
  host = document.createElement('div');
  document.body.appendChild(host);
  root = createRoot(host);
});
afterEach(() => {
  act(() => root?.unmount());
  host?.remove();
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

const render = (ctx: WriteResultContext, result: ExecResult = ok204) =>
  act(() =>
    root!.render(
      <FluentProvider theme={webLightTheme}>
        <WriteResultCard result={result} ctx={ctx} />
      </FluentProvider>,
    ),
  );

const NOTE = "ToolBox's update doesn't pass the row back";

describe('WriteResultCard after an Update / Upsert (#224)', () => {
  it.each(['update', 'upsert'] as const)(
    '%s with return=representation says why no row came back',
    (operation) => {
      render({ operation, table: 'account', recordName: 'Contoso', returnRepresentation: true });
      expect(host!.textContent).toContain(NOTE);
    },
  );

  it('points to Retrieve this record only when that button is there', () => {
    render({
      operation: 'update',
      table: 'account',
      recordId: '00000000-0000-0000-0000-0000000000a1',
      returnRepresentation: true,
    });
    expect(host!.textContent).toContain('Retrieve this record to see it.');
    // Upsert by alternate key: no id, so no button to point at.
    render({ operation: 'upsert', table: 'account', recordId: null, returnRepresentation: true });
    expect(host!.textContent).toContain(NOTE);
    expect(host!.textContent).not.toContain('Retrieve this record to see it.');
  });

  it('stays quiet when a row did come back', () => {
    render(
      { operation: 'update', table: 'account', returnRepresentation: true },
      {
        ...ok204,
        status: 200,
        statusText: 'OK',
        body: { '@odata.etag': 'W/"1"', name: 'Contoso' },
      },
    );
    expect(host!.textContent).not.toContain(NOTE);
  });

  it("doesn't mention it when return=representation wasn't asked for", () => {
    render({ operation: 'update', table: 'account', recordName: 'Contoso' });
    expect(host!.textContent).toContain('Updated');
    expect(host!.textContent).not.toContain(NOTE);
  });

  it("doesn't mention it on a failure", () => {
    render(
      { operation: 'update', table: 'account', returnRepresentation: true },
      { ...ok204, ok: false, outcome: 'error', status: 412, statusText: 'Precondition Failed' },
    );
    expect(host!.textContent).not.toContain(NOTE);
  });
});
