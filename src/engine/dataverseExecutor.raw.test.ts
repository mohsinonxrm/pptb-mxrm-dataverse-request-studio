// #225: a one-part $batch's answer mapped onto ExecResult (`fromRaw`), and
// the host errors only the $batch path raises (`mapError`).
import { describe, expect, it } from 'vitest';
import {
  BATCH_RESPONSE_UNREADABLE_CODE,
  RESPONSE_TOO_LARGE_CODE,
  fromRaw,
  mapError,
} from './dataverseExecutor';
import { BatchPathError } from '../host/batchUrl';

const errorOf = (body: unknown) =>
  (body as { error: { code?: string; message: string; rawMessage?: string } }).error;

/** How Electron wraps a rejection from ToolBox's executeBatch handler. */
const ipc = (message: string) =>
  new Error(
    `Error invoking remote method 'dataverse.executeBatch': Error: Dataverse executeBatch failed: ${message}`,
  );

describe('fromRaw', () => {
  it('204: a success with no body, the part headers and the transport', () => {
    const r = fromRaw(
      {
        status: 204,
        statusText: 'No Content',
        headers: { 'odata-entityid': 'https://x/api/data/v9.2/accounts(1)' },
      },
      12,
    );
    expect(r).toMatchObject({
      status: 204,
      statusText: 'No Content',
      ok: true,
      outcome: 'ok',
      body: null,
      ms: 12,
      transport: 'batch',
      headers: { 'odata-entityid': 'https://x/api/data/v9.2/accounts(1)' },
    });
    expect(r).not.toHaveProperty('batchHeaders');
  });

  it('201 with the returned row, and the outer x-ms headers', () => {
    const row = { accountid: '1', name: 'Contoso' };
    const r = fromRaw(
      {
        status: 201,
        headers: { etag: 'W/"5"' },
        batchHeaders: { 'x-ms-ratelimit-burst-remaining-xrm-requests': '5999' },
        body: row,
      },
      3,
    );
    expect(r).toMatchObject({
      status: 201,
      statusText: 'Created',
      ok: true,
      outcome: 'ok',
      body: row,
      headers: { etag: 'W/"5"' },
      batchHeaders: { 'x-ms-ratelimit-burst-remaining-xrm-requests': '5999' },
    });
    expect(r.bytes).toBeGreaterThan(0);
  });

  it('304 counts as success', () => {
    const r = fromRaw({ status: 304, headers: { etag: 'W/"5"' } }, 1);
    expect(r).toMatchObject({
      status: 304,
      statusText: 'Not Modified',
      ok: true,
      outcome: 'ok',
      body: null,
    });
  });

  it('412 carries the Dataverse error from the body', () => {
    const r = fromRaw(
      {
        status: 412,
        statusText: 'Precondition Failed',
        headers: {},
        body: {
          error: {
            code: '0x80060882',
            message: 'The version of the existing record doesn’t match the RowVersion property.',
          },
        },
      },
      4,
    );
    expect(r).toMatchObject({
      status: 412,
      statusText: 'Precondition Failed',
      ok: false,
      outcome: 'error',
      transport: 'batch',
    });
    expect(errorOf(r.body)).toEqual({
      code: '0x80060882',
      message: 'The version of the existing record doesn’t match the RowVersion property.',
    });
  });

  it('a part-level 429 keeps retry-after in headers', () => {
    const r = fromRaw(
      {
        status: 429,
        headers: { 'retry-after': '30' },
        body: { error: { code: '0x80072322', message: 'Number of requests exceeded the limit' } },
      },
      2,
    );
    expect(r).toMatchObject({ status: 429, statusText: 'Too Many Requests', ok: false });
    expect(r.headers['retry-after']).toBe('30');
    expect(errorOf(r.body).code).toBe('0x80072322');
  });

  it('a string body becomes the error message', () => {
    const r = fromRaw({ status: 500, headers: {}, body: '  Internal failure 0x80040216  ' }, 1);
    expect(r).toMatchObject({ status: 500, ok: false, outcome: 'error' });
    expect(errorOf(r.body)).toEqual({ code: '0x80040216', message: 'Internal failure 0x80040216' });
  });

  it('a failure with no body says the status', () => {
    const r = fromRaw({ status: 404, headers: {} }, 1);
    expect(r).toMatchObject({ status: 404, outcome: '404' });
    expect(errorOf(r.body).message).toBe('HTTP 404');
  });

  it("a 200 body that's a string is kept as the body", () => {
    expect(fromRaw({ status: 200, headers: {}, body: 'Contoso' }, 1)).toMatchObject({
      status: 200,
      ok: true,
      body: 'Contoso',
    });
  });
});

describe('mapError: $batch path errors', () => {
  it('a response over the $batch limit says so, IPC wrapper and all', () => {
    const r = mapError(ipc('Dataverse response exceeds the maximum size (16 MB)'), 9);
    const error = errorOf(r.body);
    expect(r).toMatchObject({ status: 0, ok: false, statusText: 'Response too large' });
    expect(error.code).toBe(RESPONSE_TOO_LARGE_CODE);
    expect(error.message).toMatch(/^The response is too large for the \$batch path \(16 MB\)/);
    expect(error.message).toMatch(/may have run/);
    expect(error.message).not.toMatch(/Error invoking remote method/);
    expect(error.rawMessage).toMatch(/^Error invoking remote method 'dataverse.executeBatch'/);
  });

  it("a typed call's size limit isn't called the $batch path", () => {
    const r = mapError(
      new Error(
        "Error invoking remote method 'dataverse.queryData': Error: Dataverse queryData failed: Dataverse response exceeds the maximum size (128 MB)",
      ),
      9,
    );
    const error = errorOf(r.body);
    expect(error.code).toBe(RESPONSE_TOO_LARGE_CODE);
    expect(error.message).toMatch(/too large for Power Platform ToolBox \(128 MB\)/);
    expect(error.message).not.toMatch(/\$batch/);
  });

  it('a malformed $batch response says the request may have run', () => {
    const r = mapError(ipc('Malformed Dataverse batch response: missing boundary'), 5);
    const error = errorOf(r.body);
    expect(r).toMatchObject({ status: 0, ok: false, statusText: 'Response unreadable' });
    expect(error.code).toBe(BATCH_RESPONSE_UNREADABLE_CODE);
    expect(error.message).toMatch(
      /^Power Platform ToolBox couldn't read Dataverse's answer; the request may have run\./,
    );
    expect(error.message).toMatch(/Host message: Dataverse executeBatch failed: Malformed/);
    expect(error.rawMessage).toMatch(/^Error invoking remote method/);
  });

  it("rawRequest's own empty-result error maps the same way", () => {
    const r = mapError(new Error('Malformed Dataverse batch response: no operation result'), 1);
    expect(errorOf(r.body).code).toBe(BATCH_RESPONSE_UNREADABLE_CODE);
  });

  it('a path toBatchUrl refuses is a 400: nothing was sent', () => {
    const r = mapError(new BatchPathError('the request path is empty.'), 0);
    expect(r).toMatchObject({ status: 400, statusText: 'Bad Request', ok: false });
    expect(errorOf(r.body).message).toBe('Not sent: the request path is empty.');
  });
});
