// Results sent as a one-part $batch (#225): PUT of one column, clearing one
// column and Upsert by alternate key carry Dataverse's real status, so the
// card's copy follows it, and a caption says how the request was sent.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { WriteResultCard, type WriteResultContext } from './WriteResultCard';
import type { ExecResult } from '../../engine/dataverseExecutor';
import { registerFixtureTables } from '../../test/fixtures';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let root: Root | null = null;
let host: HTMLDivElement | null = null;
beforeEach(() => {
  registerFixtureTables();
  host = document.createElement('div');
  document.body.appendChild(host);
  root = createRoot(host);
});
afterEach(() => {
  act(() => root?.unmount());
  host?.remove();
});

const ID = '00000000-0000-0000-0000-0000000000a1';
const CAPTION = 'Sent through $batch.';
const NO_ROW_NOTE = "ToolBox's update doesn't pass the row back";

const batch = (over: Partial<ExecResult>): ExecResult => ({
  status: 204,
  statusText: 'No Content',
  ms: 5,
  bytes: 0,
  ok: true,
  outcome: 'ok',
  body: null,
  headers: {},
  transport: 'batch',
  ...over,
});

const render = (ctx: WriteResultContext, result: ExecResult) =>
  act(() =>
    root!.render(
      <FluentProvider theme={webLightTheme}>
        <WriteResultCard result={result} ctx={ctx} />
      </FluentProvider>,
    ),
  );
const text = () => host!.textContent ?? '';

describe('WriteResultCard for a request sent through $batch (#225)', () => {
  it('Update by PUT: the usual success copy, and the caption', () => {
    render(
      { operation: 'update', table: 'account', recordId: ID, recordName: 'Contoso' },
      batch({}),
    );
    expect(text()).toContain('Updated Account');
    expect(text()).toContain('Saved changes to "Contoso".');
    expect(text()).toContain(CAPTION);
    expect(text()).toContain('204');
  });

  it('a stale ETag reads as a real 412', () => {
    render(
      { operation: 'update', table: 'account', recordId: ID },
      batch({
        status: 412,
        statusText: 'Precondition Failed',
        ok: false,
        outcome: 'error',
        body: { error: { code: '0x80060882', message: 'Version mismatch' } },
      }),
    );
    expect(text()).toContain('Operation rejected');
    expect(text()).toContain('the ETag is stale');
    expect(text()).toContain('412');
    expect(text()).toContain(CAPTION);
  });

  it('Delete of one column: Cleared column', () => {
    render(
      { operation: 'delete', table: 'account', recordId: ID, clearedColumn: 'description' },
      batch({}),
    );
    expect(text()).toContain('Cleared column on Account');
    expect(text()).toContain('Set description to null');
  });

  it("Upsert by alternate key, 204: says Dataverse doesn't tell create from update", () => {
    render(
      { operation: 'upsert', table: 'account', recordId: null, returnRepresentation: false },
      batch({}),
    );
    expect(text()).toContain('Upsert completed on Account');
    expect(text()).toContain('Dataverse answers 204 whether it created or updated the row.');
    expect(text()).not.toContain("PPTB host doesn't surface");
  });

  it('Upsert by alternate key, 201 with the row: created, and the row is shown', () => {
    render(
      { operation: 'upsert', table: 'account', recordId: null, returnRepresentation: true },
      batch({
        status: 201,
        statusText: 'Created',
        body: { '@odata.etag': 'W/"9"', accountid: ID, name: 'Fabrikam' },
      }),
    );
    expect(text()).toContain('Upsert: created new Account');
    expect(text()).toContain('Returned representation');
    expect(text()).not.toContain(NO_ROW_NOTE);
  });

  it('Upsert by alternate key, 200 with the row: updated', () => {
    render(
      { operation: 'upsert', table: 'account', recordId: null, returnRepresentation: true },
      batch({ status: 200, statusText: 'OK', body: { '@odata.etag': 'W/"9"', accountid: ID } }),
    );
    expect(text()).toContain('Upsert: updated Account');
    expect(text()).not.toContain(NO_ROW_NOTE);
  });

  it('Upsert by alternate key, 404 (update-only, no such row): names the key, not an id', () => {
    render(
      { operation: 'upsert', table: 'account', recordId: null },
      batch({
        status: 404,
        statusText: 'Not Found',
        ok: false,
        outcome: '404',
        body: { error: { code: '0x80040217', message: 'Does Not Exist' } },
      }),
    );
    expect(text()).toContain('Record not found');
    expect(text()).toContain('No Account record matched the alternate key.');
  });

  it('a typed result has no caption', () => {
    const { transport: _t, ...typed } = batch({});
    render({ operation: 'update', table: 'account', recordId: ID }, typed);
    expect(text()).not.toContain(CAPTION);
  });
});
