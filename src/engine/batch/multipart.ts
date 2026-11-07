// multipart/mixed $batch body writer + response reader (#9).
//
// Format: OData v4 §11.7 as profiled by Dataverse —
// https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/execute-batch-operations-using-web-api
//
//   • every line ends in CRLF
//   • each part: `Content-Type: application/http`, `Content-Transfer-Encoding:
//     binary`, `Content-ID: n`, blank line, the HTTP request, blank line, body
//   • a change set is a nested multipart/mixed part with its own boundary
//   • bodies go out as `Content-Type: application/json; type=entry`
//
// The writer mirrors what PPTB's `encodeDataverseBatch` sends (desktop-app
// src/main/utils/dataverseBatch.ts), so the Code tab preview is the request
// that goes on the wire, apart from the random boundaries. The reader parses
// a raw multipart response (nested change sets included) into one entry per
// operation — used to inspect a response pasted from another client.

import { API_BASE, type WireRequest } from './batchModel';

export const CRLF = '\r\n';

export interface EncodablePart {
  contentId: string;
  wire: WireRequest;
}

export interface EncodableGroup {
  /** True → wrap the parts in a change set. */
  atomic: boolean;
  parts: EncodablePart[];
}

export interface EncodedBatch {
  boundary: string;
  contentType: string;
  body: string;
}

/** `/api/data/v9.2/<path>`, except a leading `$n` reference stays bare (OData §11.7.3.1). */
export function wireUrl(path: string): string {
  return /^\$\d+/.test(path) ? path : `${API_BASE}/${path}`;
}

function hasHeader(headers: Record<string, string>, name: string): boolean {
  const lower = name.toLowerCase();
  return Object.keys(headers).some((h) => h.toLowerCase() === lower);
}

export function encodePart({ contentId, wire }: EncodablePart): string {
  const lines = [
    'Content-Type: application/http',
    'Content-Transfer-Encoding: binary',
    `Content-ID: ${contentId}`,
    '',
    `${wire.method} ${wireUrl(wire.path)} HTTP/1.1`,
  ];
  for (const [name, value] of Object.entries(wire.headers)) lines.push(`${name}: ${value}`);
  if (wire.body !== undefined && !hasHeader(wire.headers, 'Content-Type')) {
    lines.push('Content-Type: application/json; type=entry');
  }
  lines.push('', wire.body ?? '');
  return lines.join(CRLF);
}

let counter = 0;
/** Unique-enough boundary token. Not a security boundary — just has to not appear in the body. */
export function makeBoundary(prefix: 'batch' | 'changeset'): string {
  counter += 1;
  const rand =
    typeof crypto !== 'undefined' && 'randomUUID' in crypto
      ? crypto.randomUUID()
      : `${Date.now().toString(16)}-${counter.toString(16)}-${Math.random().toString(16).slice(2)}`;
  return `${prefix}_${rand}`;
}

/**
 * Encode groups into one $batch body. `boundaries` lets tests pin the
 * tokens; by default they're random.
 */
export function encodeBatch(
  groups: EncodableGroup[],
  boundaries: { batch?: string; changeset?: (index: number) => string } = {},
): EncodedBatch {
  const boundary = boundaries.batch ?? makeBoundary('batch');
  let changesetIndex = 0;
  let body = '';
  for (const group of groups) {
    if (group.parts.length === 0) continue;
    if (group.atomic) {
      const cs = boundaries.changeset?.(changesetIndex) ?? makeBoundary('changeset');
      changesetIndex += 1;
      body += `--${boundary}${CRLF}Content-Type: multipart/mixed; boundary=${cs}${CRLF}${CRLF}`;
      for (const part of group.parts) body += `--${cs}${CRLF}${encodePart(part)}${CRLF}`;
      body += `--${cs}--${CRLF}`;
    } else {
      for (const part of group.parts) body += `--${boundary}${CRLF}${encodePart(part)}${CRLF}`;
    }
  }
  body += `--${boundary}--${CRLF}`;
  return { boundary, contentType: `multipart/mixed; boundary=${boundary}`, body };
}

// ── Response reader ────────────────────────────────────────────────────

export interface ParsedBatchResponse {
  contentId?: string;
  status: number;
  statusText?: string;
  headers: Record<string, string>;
  body?: unknown;
  /** Index of the change set this response came from (0-based), if any. */
  changeset?: number;
}

const BOUNDARY_RE = /boundary=(?:"([^"]+)"|([^;\s]+))/i;

/** Pull the boundary out of a Content-Type header value, or out of the body's first line. */
export function findBoundary(contentType: string, body: string): string | null {
  const m = BOUNDARY_RE.exec(contentType);
  if (m) return m[1] || m[2];
  const first = /^\s*--([^\r\n]+)/.exec(body);
  return first ? first[1].trim() : null;
}

function parseHeaderBlock(block: string): Record<string, string> {
  const headers: Record<string, string> = {};
  for (const line of block.split(/\r?\n/)) {
    const i = line.indexOf(':');
    if (i <= 0) continue;
    const name = line.slice(0, i).trim().toLowerCase();
    const value = line.slice(i + 1).trim();
    headers[name] = headers[name] ? `${headers[name]}, ${value}` : value;
  }
  return headers;
}

function splitParts(body: string, boundary: string): string[] {
  const delimiter = `--${boundary}`;
  const parts: string[] = [];
  let cursor = body.indexOf(delimiter);
  if (cursor < 0) throw new Error(`Boundary "${boundary}" not found.`);
  while (cursor >= 0) {
    cursor += delimiter.length;
    if (body.startsWith('--', cursor)) return parts;
    const next = body.indexOf(delimiter, cursor);
    if (next < 0) throw new Error('The body is missing its closing boundary.');
    parts.push(
      body
        .slice(cursor, next)
        .replace(/^\r?\n/, '')
        .replace(/\r?\n$/, ''),
    );
    cursor = next;
  }
  return parts;
}

function splitHead(text: string): [string, string] {
  const m = /\r?\n\r?\n/.exec(text);
  if (!m) return [text, ''];
  return [text.slice(0, m.index), text.slice(m.index + m[0].length)];
}

// ── Request reader (cURL import, #271) ─────────────────────────────────

export interface ParsedBatchRequest {
  method: string;
  /** As written on the request line: absolute, `/api/data/…`, relative or `$n`. */
  url: string;
  /** Lower-cased names. */
  headers: Record<string, string>;
  body: string;
  contentId?: string;
}

export interface ParsedBatchRequestGroup {
  /** True for a change set. */
  atomic: boolean;
  requests: ParsedBatchRequest[];
}

const BATCH_METHODS = new Set(['GET', 'POST', 'PATCH', 'PUT', 'DELETE']);

/**
 * Parse a multipart/mixed $batch request body into its requests, grouped as
 * sent: each change set is one atomic group, each request outside one is a
 * group of its own.
 */
export function parseBatchRequest(body: string, contentType = ''): ParsedBatchRequestGroup[] {
  const boundary = findBoundary(contentType, body);
  if (!boundary) throw new Error('No multipart boundary in the Content-Type or the body.');
  const readRequest = (part: string): ParsedBatchRequest => {
    const [mimeHead, payload] = splitHead(part);
    const mime = parseHeaderBlock(mimeHead);
    const [httpHead, rawBody] = splitHead(payload.replace(/^\s+/, ''));
    const [requestLine, ...headerLines] = httpHead.split(/\r?\n/);
    const m = /^([A-Z]+)\s+(\S+)(?:\s+HTTP\/1\.[01])?$/i.exec(requestLine.trim());
    if (!m) throw new Error(`Malformed request line: ${requestLine.trim().slice(0, 80)}`);
    const method = m[1].toUpperCase();
    if (!BATCH_METHODS.has(method)) throw new Error(`${method} isn't allowed in a $batch.`);
    const req: ParsedBatchRequest = {
      method,
      url: m[2],
      headers: parseHeaderBlock(headerLines.join('\n')),
      body: rawBody.trim(),
    };
    if (mime['content-id']) req.contentId = mime['content-id'];
    return req;
  };
  return splitParts(body, boundary).map((part) => {
    const [mimeHead, payload] = splitHead(part);
    const nested = BOUNDARY_RE.exec(parseHeaderBlock(mimeHead)['content-type'] ?? '');
    if (nested) {
      return {
        atomic: true,
        requests: splitParts(payload, nested[1] || nested[2]).map(readRequest),
      };
    }
    return { atomic: false, requests: [readRequest(part)] };
  });
}

/**
 * Parse a raw multipart/mixed $batch response into one entry per operation,
 * in order. Change-set responses are flattened and tagged with their index.
 */
export function parseBatchResponse(body: string, contentType = ''): ParsedBatchResponse[] {
  const boundary = findBoundary(contentType, body);
  if (!boundary) throw new Error('No multipart boundary in the Content-Type or the body.');
  let changesetIndex = 0;
  const walk = (text: string, b: string, changeset?: number): ParsedBatchResponse[] =>
    splitParts(text, b).flatMap((part) => {
      const [mimeHead, payload] = splitHead(part);
      const mime = parseHeaderBlock(mimeHead);
      const nested = BOUNDARY_RE.exec(mime['content-type'] ?? '');
      if (nested) {
        const index = changesetIndex++;
        return walk(payload, nested[1] || nested[2], index);
      }
      const statusAt = payload.search(/HTTP\/1\.[01]\s+\d{3}/i);
      if (statusAt < 0) throw new Error('A response part has no HTTP status line.');
      const [httpHead, rawBody] = splitHead(payload.slice(statusAt));
      const [statusLine, ...headerLines] = httpHead.split(/\r?\n/);
      const sm = /^HTTP\/1\.[01]\s+(\d{3})(?:\s+(.*))?$/i.exec(statusLine.trim());
      if (!sm) throw new Error(`Malformed status line: ${statusLine}`);
      const text = rawBody.trim();
      let parsed: unknown = undefined;
      if (text) {
        try {
          parsed = JSON.parse(text);
        } catch {
          parsed = text;
        }
      }
      const result: ParsedBatchResponse = {
        status: Number(sm[1]),
        statusText: sm[2]?.trim() || undefined,
        headers: parseHeaderBlock(headerLines.join('\n')),
      };
      if (mime['content-id']) result.contentId = mime['content-id'];
      if (parsed !== undefined) result.body = parsed;
      if (changeset !== undefined) result.changeset = changeset;
      return [result];
    });
  return walk(body, boundary);
}
