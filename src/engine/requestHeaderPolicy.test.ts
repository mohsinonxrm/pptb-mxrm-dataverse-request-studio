// #109: which headers the read, relate and Execute Function modes send.
import { describe, expect, it } from 'vitest';
import {
  allowedHeaders,
  readPrefer,
  readRequestHeaders,
  unsentHeaderNames,
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

  it('Execute Function: impersonation only', () => {
    expect(allowedHeaders('function', PANE)).toEqual({ MSCRMCallerID: ID, CallerObjectId: ID });
  });

  it('matches names case-insensitively and skips the old If-None-Match: null default', () => {
    expect(allowedHeaders('read', { mscrmcallerid: ID, 'If-None-Match': 'null' })).toEqual({
      mscrmcallerid: ID,
    });
    expect(allowedHeaders('read', { Accept: 'application/json' })).toBeUndefined();
  });
});

describe('unsentHeaderNames', () => {
  it('lists enabled rows Execute leaves to the Code tab, not host-owned or credential ones', () => {
    expect(unsentHeaderNames('function', PANE)).toEqual([
      'If-None-Match',
      'Consistency',
      'MSCRM.BypassBusinessLogicExecution',
      'X-Custom',
    ]);
  });
});

describe('readPrefer', () => {
  it('nothing beyond the host default: no Prefer, no prompt', () => {
    expect(readPrefer(null)).toBeUndefined();
    expect(
      readPrefer(null, 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"'),
    ).toBeUndefined();
  });

  it('keeps every annotation first and the page size last', () => {
    expect(readPrefer(25, 'odata.track-changes')).toBe(
      'odata.include-annotations="*",odata.track-changes,odata.maxpagesize=25',
    );
  });

  it('splits on commas outside quotes and drops annotation, page size and return tokens', () => {
    expect(
      readPrefer(
        null,
        'odata.include-annotations="a,b", return=representation, odata.maxpagesize=9, respond-async',
      ),
    ).toBe('odata.include-annotations="*",respond-async');
  });
});

describe('readRequestHeaders', () => {
  it('merges the allow-listed rows with the read Prefer', () => {
    expect(readRequestHeaders({ MSCRMCallerID: ID, prefer: 'odata.track-changes' }, 10)).toEqual({
      MSCRMCallerID: ID,
      Prefer: 'odata.include-annotations="*",odata.track-changes,odata.maxpagesize=10',
    });
    expect(readRequestHeaders({ Accept: 'application/json' }, null)).toBeUndefined();
  });
});
