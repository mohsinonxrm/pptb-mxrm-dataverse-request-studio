// #109: which headers the read, relate, Execute Function and Bulk Delete
// modes send.
import { describe, expect, it } from 'vitest';
import {
  allowedHeaders,
  impersonationProblem,
  readPrefer,
  readRequestHeaders,
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
