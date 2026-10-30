// #222: a file name outside ASCII can't travel in the x-ms-file-name
// header (Dataverse doesn't support it there, and fetch / XHR throw), so it
// goes in the query parameter instead.
import { beforeAll, describe, expect, it } from 'vitest';
import {
  buildFileChunkedUpload,
  buildFileSingleUpload,
  isHeaderSafeFileName,
} from './binaryBuilders';
import { __registerLiveTables, type TableMeta } from '../mock/metadata';
import type { ManageFileState } from '../state/binaryState';

const ID = '00000000-0000-0000-0000-0000000000a1';

beforeAll(() => {
  __registerLiveTables([
    {
      logicalName: 'account',
      entitySetName: 'accounts',
      displayName: 'Account',
      primaryKey: 'accountid',
      primaryName: 'name',
      columns: [],
      navigationProperties: [],
    } as TableMeta,
  ]);
});

const file = (fileName: string, uploadMethod: ManageFileState['uploadMethod']) =>
  ({
    table: 'account',
    recordId: ID,
    fileColumn: 'sample_filecolumn',
    operation: 'upload',
    uploadMethod,
    source: { kind: 'file' },
    fileName,
    fileSize: 10 * 1024 * 1024,
    mimeType: 'application/pdf',
    bodyBase64: '',
  }) as unknown as ManageFileState;

/** What a browser would do with the step's headers. */
const sendable = (headers: Record<string, string> = {}) =>
  // Latin-1 only, as fetch / XHR require of a header value.
  // eslint-disable-next-line no-control-regex
  Object.values(headers).every((v) => /^[\x00-\xff]*$/.test(v));

describe('isHeaderSafeFileName', () => {
  it('printable ASCII only', () => {
    expect(isHeaderSafeFileName('4094kb.txt')).toBe(true);
    expect(isHeaderSafeFileName('My report (1).pdf')).toBe(true);
    expect(isHeaderSafeFileName('Résumé.docx')).toBe(false);
    expect(isHeaderSafeFileName('报告.pdf')).toBe(false);
  });
});

describe('single-request upload', () => {
  it('an ASCII name stays in the header, as Microsoft documents', () => {
    const [step] = buildFileSingleUpload(file('4094kb.txt', { kind: 'single-request' } as never));
    expect(step.relativeUrl).toBe(`/api/data/v9.2/accounts(${ID})/sample_filecolumn`);
    expect(step.headers?.['x-ms-file-name']).toBe('4094kb.txt');
  });

  it('a non-ASCII name goes in the query parameter, with no header', () => {
    const [step] = buildFileSingleUpload(
      file('报告 Résumé✓.pdf', { kind: 'single-request' } as never),
    );
    expect(step.relativeUrl).toBe(
      `/api/data/v9.2/accounts(${ID})/sample_filecolumn?x-ms-file-name=${encodeURIComponent('报告 Résumé✓.pdf')}`,
    );
    expect(step.headers?.['x-ms-file-name']).toBeUndefined();
    expect(sendable(step.headers)).toBe(true);
    expect(step.detail).toMatch(/outside ASCII/);
  });
});

describe('chunked upload', () => {
  const chunked = { kind: 'chunked-patch', chunkSize: 4 * 1024 * 1024 } as never;

  it('ASCII: the init query parameter and a header on every chunk', () => {
    const steps = buildFileChunkedUpload(file('25mb.pdf', chunked));
    expect(steps[0].relativeUrl).toMatch(/\?x-ms-file-name=25mb\.pdf$/);
    expect(steps.slice(1).every((s) => s.headers?.['x-ms-file-name'] === '25mb.pdf')).toBe(true);
  });

  it('non-ASCII: the init query parameter only; no chunk carries the header', () => {
    const steps = buildFileChunkedUpload(file('Ünïcødé.pdf', chunked));
    expect(steps[0].relativeUrl).toMatch(
      new RegExp(`\\?x-ms-file-name=${encodeURIComponent('Ünïcødé.pdf')}$`),
    );
    expect(steps.slice(1).every((s) => s.headers?.['x-ms-file-name'] === undefined)).toBe(true);
    expect(steps.every((s) => sendable(s.headers))).toBe(true);
  });
});
