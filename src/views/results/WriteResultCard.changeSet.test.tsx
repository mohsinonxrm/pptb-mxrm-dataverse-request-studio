// All or nothing (#224): what the result card says about a relate change set.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { WriteResultCard, type WriteResultContext } from './WriteResultCard';
import type { ExecResult, RelateTargetResult } from '../../engine/dataverseExecutor';

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

const ctx: WriteResultContext = {
  operation: 'associate',
  table: 'account',
  recordId: '00000000-0000-0000-0000-00000000c0de',
  recordName: 'Contoso',
  navProperty: 'contact_customer_accounts',
  targetCount: 2,
  targetNames: [],
  targetNameById: {},
};
const T1 = '00000000-0000-0000-0000-0000000000a1';
const T2 = '00000000-0000-0000-0000-0000000000a2';

const result = (
  status: number,
  targets: RelateTargetResult[],
  error?: { code?: string; message: string },
): ExecResult => ({
  status,
  statusText: status === 204 ? 'No Content' : 'Error',
  ms: 5,
  bytes: 0,
  ok: status === 204,
  outcome: status === 204 ? 'ok' : 'error',
  body: error ? { error, targets } : { targets },
  headers: {},
  targets,
  changeSet: true,
});

const render = (r: ExecResult) =>
  act(() =>
    root!.render(
      <FluentProvider theme={webLightTheme}>
        <WriteResultCard result={r} ctx={ctx} />
      </FluentProvider>,
    ),
  );

describe('WriteResultCard, All or nothing (#224)', () => {
  it('says a successful run went as one change set', () => {
    render(
      result(204, [
        { targetId: T1, state: 'done' },
        { targetId: T2, state: 'done' },
      ]),
    );
    expect(host!.textContent).toContain('Associated 2 records to');
    expect(host!.textContent).toContain('Sent as one change set.');
  });

  it('names the failed target and marks the other not applied', () => {
    render(
      result(
        404,
        [
          { targetId: T1, state: 'rolled-back' },
          {
            targetId: T2,
            state: 'failed',
            status: 404,
            statusText: 'Not Found',
            error: { code: '0x80040217', message: 'contact Does Not Exist' },
          },
        ],
        { code: '0x80040217', message: 'contact Does Not Exist' },
      ),
    );
    const text = host!.textContent!;
    expect(text).toContain('Associate failed — nothing changed');
    expect(text).toContain('One failed, so Dataverse rolled back the other 1');
    expect(text).toContain('Not applied');
    expect(text).toContain('Failed · 404 Not Found');
    expect(text).toContain('contact Does Not Exist');
  });

  it("shows the error when it can't be tied to a target", () => {
    render(
      result(
        400,
        [
          { targetId: T1, state: 'rolled-back' },
          { targetId: T2, state: 'rolled-back' },
        ],
        { message: 'The change set was rejected' },
      ),
    );
    const text = host!.textContent!;
    expect(text).toContain('Associate failed — change set not applied');
    expect(text).toContain('The change set of 2 requests failed: The change set was rejected.');
    expect(text).not.toContain('rolled back the other');
  });
});
