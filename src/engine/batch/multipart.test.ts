import { describe, expect, it } from 'vitest';
import { encodeBatch, encodePart, parseBatchResponse, wireUrl, CRLF } from './multipart';
import { prettyJson, type WireRequest } from './batchModel';

const create = (body: object): WireRequest => ({
  method: 'POST',
  path: 'tasks',
  headers: {},
  body: prettyJson(body),
});

describe('encodePart', () => {
  it('writes the MIME headers, the request line and a typed JSON body', () => {
    const part = encodePart({ contentId: '1', wire: create({ subject: 'x' }) });
    expect(part.split(CRLF)).toEqual([
      'Content-Type: application/http',
      'Content-Transfer-Encoding: binary',
      'Content-ID: 1',
      '',
      'POST /api/data/v9.2/tasks HTTP/1.1',
      'Content-Type: application/json; type=entry',
      '',
      '{',
      '  "subject": "x"',
      '}',
    ]);
  });

  it('keeps a user Content-Type and per-request headers', () => {
    const part = encodePart({
      contentId: '2',
      wire: {
        method: 'PATCH',
        path: 'accounts(1)',
        headers: { 'If-Match': '*', 'Content-Type': 'application/json' },
        body: '{}',
      },
    });
    expect(part).toContain(`PATCH /api/data/v9.2/accounts(1) HTTP/1.1${CRLF}If-Match: *`);
    expect(part).not.toContain('type=entry');
  });

  it('leaves Content-ID references in the URL bare', () => {
    expect(wireUrl('$1/lastname')).toBe('$1/lastname');
    expect(wireUrl('accounts')).toBe('/api/data/v9.2/accounts');
  });

  it('writes no body for a GET', () => {
    const part = encodePart({
      contentId: '3',
      wire: { method: 'GET', path: 'accounts?$top=1', headers: {} },
    });
    expect(part.endsWith(`GET /api/data/v9.2/accounts?$top=1 HTTP/1.1${CRLF}${CRLF}`)).toBe(true);
  });
});

describe('encodeBatch', () => {
  it('nests change sets and closes every boundary', () => {
    const { body, contentType } = encodeBatch(
      [
        {
          atomic: true,
          parts: [
            { contentId: '1', wire: create({ a: 1 }) },
            { contentId: '2', wire: create({ b: 2 }) },
          ],
        },
        {
          atomic: false,
          parts: [{ contentId: '3', wire: { method: 'GET', path: 'tasks', headers: {} } }],
        },
      ],
      { batch: 'batch_B', changeset: (i) => `changeset_C${i}` },
    );
    expect(contentType).toBe('multipart/mixed; boundary=batch_B');
    expect(
      body.startsWith(`--batch_B${CRLF}Content-Type: multipart/mixed; boundary=changeset_C0`),
    ).toBe(true);
    expect(body.split('--changeset_C0' + CRLF).length - 1).toBe(2);
    expect(body).toContain(`--changeset_C0--${CRLF}--batch_B${CRLF}Content-Type: application/http`);
    expect(body.endsWith(`--batch_B--${CRLF}`)).toBe(true);
    // CRLF only — no bare LF anywhere, JSON bodies included.
    expect(/[^\r]\n/.test(body)).toBe(false);
  });

  it('skips empty groups', () => {
    const { body } = encodeBatch([{ atomic: true, parts: [] }], { batch: 'batch_X' });
    expect(body).toBe(`--batch_X--${CRLF}`);
  });
});

// Response samples from the Microsoft Learn article (CRLF restored).
const crlf = (s: string) => s.replace(/\r?\n/g, CRLF);

const CHANGESET_RESPONSE = crlf(`--batchresponse_f27e
Content-Type: multipart/mixed; boundary=changesetresponse_64cc

--changesetresponse_64cc
Content-Type: application/http
Content-Transfer-Encoding: binary
Content-ID: 1

HTTP/1.1 204 No Content
OData-Version: 4.0
OData-EntityId: https://org.crm.dynamics.com/api/data/v9.2/tasks(33dd33dd-ee44-ff55-aa66-77bb77bb77bb)

--changesetresponse_64cc
Content-Type: application/http
Content-Transfer-Encoding: binary
Content-ID: 2

HTTP/1.1 204 No Content
OData-Version: 4.0

--changesetresponse_64cc--
--batchresponse_f27e
Content-Type: application/http
Content-Transfer-Encoding: binary

HTTP/1.1 200 OK
Content-Type: application/json; odata.metadata=minimal
OData-Version: 4.0

{"value":[{"subject":"Task 1 in batch"}]}
--batchresponse_f27e--
`);

const ERROR_RESPONSE = crlf(`--batchresponse_156d
Content-Type: application/http
Content-Transfer-Encoding: binary

HTTP/1.1 400 Bad Request
Content-Type: application/json; odata.metadata=minimal
OData-Version: 4.0

{"error":{"code":"0x80044331","message":"A validation error occurred."}}
--batchresponse_156d--
`);

describe('parseBatchResponse', () => {
  it('flattens change-set responses and keeps Content-IDs', () => {
    const res = parseBatchResponse(
      CHANGESET_RESPONSE,
      'multipart/mixed; boundary=batchresponse_f27e',
    );
    expect(res).toHaveLength(3);
    expect(res[0]).toMatchObject({
      contentId: '1',
      status: 204,
      statusText: 'No Content',
      changeset: 0,
    });
    expect(res[0].headers['odata-entityid']).toContain('tasks(33dd33dd');
    expect(res[1]).toMatchObject({ contentId: '2', status: 204, changeset: 0 });
    expect(res[2].contentId).toBeUndefined();
    expect(res[2].changeset).toBeUndefined();
    expect(res[2].status).toBe(200);
    expect(res[2].body).toEqual({ value: [{ subject: 'Task 1 in batch' }] });
  });

  it('reads an error part and finds the boundary from the body when no header is given', () => {
    const res = parseBatchResponse(ERROR_RESPONSE);
    expect(res).toEqual([
      {
        status: 400,
        statusText: 'Bad Request',
        headers: {
          'content-type': 'application/json; odata.metadata=minimal',
          'odata-version': '4.0',
        },
        body: { error: { code: '0x80044331', message: 'A validation error occurred.' } },
      },
    ]);
  });

  it('accepts LF-only text pasted from an editor', () => {
    const res = parseBatchResponse(CHANGESET_RESPONSE.replace(/\r\n/g, '\n'));
    expect(res.map((r) => r.status)).toEqual([204, 204, 200]);
  });

  it('keeps a non-JSON body as text', () => {
    const res = parseBatchResponse(
      crlf(`--b
Content-Type: application/http

HTTP/1.1 500 Internal Server Error

oops
--b--
`),
    );
    expect(res[0].body).toBe('oops');
  });

  it('throws on a truncated response', () => {
    expect(() => parseBatchResponse(ERROR_RESPONSE.replace('--batchresponse_156d--', ''))).toThrow(
      /closing boundary/,
    );
  });

  it('round-trips what encodeBatch writes as far as structure goes', () => {
    const { body, contentType } = encodeBatch([
      {
        atomic: false,
        parts: [{ contentId: '1', wire: { method: 'GET', path: 'x', headers: {} } }],
      },
    ]);
    // A request isn't a response, but the walker must still find exactly one part.
    expect(() => parseBatchResponse(body, contentType)).toThrow(/status line/);
  });
});
