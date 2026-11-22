// #109: which headers the read, relate, Execute Function and Bulk Delete
// modes send.
import { describe, expect, it } from 'vitest';
import {
  allowedHeaders,
  impersonationProblem,
  readPrefer,
  readRequestHeaders,
  toRawPartHeaders,
} from './requestHeaderPolicy';

const ID = '00000000-0000-0000-0000-0000000000c1';
const PANE = {
  Accept: 'application/json',
  'OData-Version': '4.0',
  MSCRMCallerID: ID,
  CallerObjectId: ID,
  'If-None-Match': 'W/"123"',
  Consistency: 'Strong',
  'MSCRM.BypassBusinessLogicExecution': 'CustomSync',
  'X-Custom': 'x',
  Cookie: 'secret',
};

describe('allowedHeaders', () => {
  it('reads: impersonation, If-None-Match and Consistency', () => {
    expect(allowedHeaders('read', PANE)).toEqual({
      MSCRMCallerID: ID,
      CallerObjectId: ID,
      'If-None-Match': 'W/"123"',
      Consistency: 'Strong',
    });
  });

  it('relate: impersonation and bypass', () => {
    expect(allowedHeaders('relate', PANE)).toEqual({
      MSCRMCallerID: ID,
      CallerObjectId: ID,
      'MSCRM.BypassBusinessLogicExecution': 'CustomSync',
    });
  });

  it('Execute Function and Bulk Delete: impersonation only', () => {
    expect(allowedHeaders('impersonation', PANE)).toEqual({
      MSCRMCallerID: ID,
      CallerObjectId: ID,
    });
  });

  it('matches names case-insensitively and skips the old If-None-Match: null default', () => {
    expect(allowedHeaders('read', { mscrmcallerid: ID, 'If-None-Match': 'null' })).toEqual({
      mscrmcallerid: ID,
    });
    expect(allowedHeaders('read', { Accept: 'application/json' })).toBeUndefined();
  });
});

describe('impersonationProblem', () => {
  it('passes one header with a user id, or none', () => {
    expect(impersonationProblem({ MSCRMCallerID: ID })).toBeNull();
    expect(impersonationProblem({ CallerObjectId: ID })).toBeNull();
    expect(impersonationProblem({ Accept: 'application/json' })).toBeNull();
    expect(impersonationProblem(undefined)).toBeNull();
  });

  it('refuses both at once, the placeholder, and a value that is not a GUID', () => {
    expect(impersonationProblem({ MSCRMCallerID: ID, CallerObjectId: ID })).toBe(
      'Turn on one impersonation header, not both (MSCRMCallerID and CallerObjectId).',
    );
    expect(impersonationProblem({ MSCRMCallerID: '00000000-0000-0000-0000-000000000000' })).toMatch(
      /^MSCRMCallerID needs the id of the user to run as/,
    );
    expect(impersonationProblem({ CallerObjectId: 'someone@contoso.com' })).toMatch(
      /^CallerObjectId needs the id/,
    );
  });
});

describe('readPrefer', () => {
  it('only for a page size: annotations first, page size last', () => {
    expect(readPrefer(null)).toBeUndefined();
    expect(readPrefer(0)).toBeUndefined();
    expect(readPrefer(25)).toBe('odata.include-annotations="*",odata.maxpagesize=25');
  });
});

describe('readRequestHeaders', () => {
  it('merges the allow-listed rows with the page-size Prefer; a Prefer row stays in the Code tab', () => {
    expect(readRequestHeaders({ MSCRMCallerID: ID, Prefer: 'odata.track-changes' }, 10)).toEqual({
      MSCRMCallerID: ID,
      Prefer: 'odata.include-annotations="*",odata.maxpagesize=10',
    });
    expect(readRequestHeaders({ Prefer: 'odata.track-changes' }, null)).toBeUndefined();
    expect(readRequestHeaders({ Accept: 'application/json' }, null)).toBeUndefined();
  });
});

// #225: the headers a one-part $batch carries. Nothing in a part is
// host-owned, so it's the Code tab's headers minus credentials and the
// values Dataverse assumes anyway.
describe('toRawPartHeaders', () => {
  it('drops Authorization, credential headers and Content-Length', () => {
    expect(
      toRawPartHeaders({
        Authorization: 'Bearer x',
        'Proxy-Authorization': 'Basic y',
        cookie: 'a=b',
        'Content-Length': '12',
        'If-Match': '*',
      }),
    ).toEqual({ 'If-Match': '*' });
  });

  it('drops Accept / OData-Version / OData-MaxVersion at their defaults, in any case', () => {
    expect(
      toRawPartHeaders({
        accept: 'Application/JSON',
        'OData-Version': '4.0',
        'odata-maxversion': ' 4.0 ',
      }),
    ).toEqual({});
  });

  it('keeps them when they differ from the defaults', () => {
    expect(
      toRawPartHeaders({
        Accept: 'application/json;odata.metadata=full',
        'OData-MaxVersion': '4.01',
      }),
    ).toEqual({ Accept: 'application/json;odata.metadata=full', 'OData-MaxVersion': '4.01' });
  });

  it('is {} when only the defaults are left', () => {
    expect(
      toRawPartHeaders({
        Accept: 'application/json',
        'Content-Type': 'application/json; charset=utf-8',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
      }),
    ).toEqual({});
    expect(toRawPartHeaders({})).toEqual({});
    expect(toRawPartHeaders(undefined)).toEqual({});
  });

  it('keeps Prefer, preconditions, bypass, impersonation and custom headers', () => {
    const headers = {
      Prefer: 'return=representation',
      'If-Match': 'W/"123"',
      'If-None-Match': '*',
      'MSCRM.BypassBusinessLogicExecution': 'CustomSync',
      'MSCRM.SuppressDuplicateDetection': 'false',
      MSCRMCallerID: ID,
      Consistency: 'Strong',
      'X-Custom': 'x',
    };
    expect(toRawPartHeaders({ ...headers, Accept: 'application/json' })).toEqual(headers);
  });

  it('drops a JSON Content-Type and keeps any other', () => {
    expect(toRawPartHeaders({ 'Content-Type': 'application/json' })).toEqual({});
    expect(toRawPartHeaders({ 'content-type': 'Application/Json; type=entry' })).toEqual({});
    expect(toRawPartHeaders({ 'Content-Type': 'text/plain' })).toEqual({
      'Content-Type': 'text/plain',
    });
    expect(toRawPartHeaders({ 'Content-Type': 'application/jsonx' })).toEqual({
      'Content-Type': 'application/jsonx',
    });
  });

  it('de-duplicates names case-insensitively, the later row winning', () => {
    expect(toRawPartHeaders({ prefer: 'return=minimal', Prefer: 'return=representation' })).toEqual(
      { Prefer: 'return=representation' },
    );
    // A later row at the default value removes an earlier custom one.
    expect(toRawPartHeaders({ Accept: 'text/plain', accept: 'application/json' })).toEqual({});
  });
});
