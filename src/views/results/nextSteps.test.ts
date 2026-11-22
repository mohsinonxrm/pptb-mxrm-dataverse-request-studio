import { describe, expect, it } from 'vitest';
import {
  describeMissingPrivilege,
  errorReport,
  missingPrivilege,
  notFoundSegment,
  replacePathSegment,
  suggestEntitySet,
} from './nextSteps';
import type { EntityListItem } from '../../host/metadataProvider';
import type { ExecResult } from '../../engine/dataverseExecutor';

const entity = (logicalName: string, entitySetName: string): EntityListItem => ({
  logicalName,
  entitySetName,
  displayName: logicalName,
  metadataId: logicalName,
});
const ENTITIES = [
  entity('account', 'accounts'),
  entity('contact', 'contacts'),
  entity('systemuser', 'systemusers'),
  entity('team', 'teams'),
  entity('cso_widget', 'cso_widgets'),
];

describe('missingPrivilege (#291)', () => {
  it('reads the privilege, its access right and the table', () => {
    const p = missingPrivilege(
      "Principal user (Id=1, type=8, roleCount=1) is missing prvReadAccount privilege (Id=2) on OTC=1 for entity 'account'.",
    );
    expect(p).toEqual({ privilege: 'prvReadAccount', access: 'Read', entity: 'account' });
    expect(describeMissingPrivilege(p!)).toBe(
      'Your user is missing the prvReadAccount privilege (Read access on account). A Dataverse administrator can add it to one of your security roles.',
    );
  });

  it('reads Append To before Append', () => {
    expect(missingPrivilege('is missing prvAppendToContact privilege')?.access).toBe('Append To');
    expect(missingPrivilege('is missing prvAppendContact privilege')?.access).toBe('Append');
  });

  it('falls back to AccessRights, and explains a record-level check', () => {
    const p = missingPrivilege(
      'SecLib::AccessCheckEx failed. Returned hr = -2147187962, ObjectTypeCode: 1, AccessRights: WriteAccess',
    );
    expect(p).toEqual({ access: 'Write' });
    // Not "add the privilege": the user may hold it at too narrow a depth.
    expect(describeMissingPrivilege(p!)).toBe(
      "Your user doesn't have Write access to this record. A security role may grant it at too narrow a depth (for example User rather than Business Unit), or the record isn't owned by or shared with you.",
    );
  });

  it('returns null when the message names nothing', () => {
    expect(missingPrivilege('Forbidden')).toBeNull();
    expect(missingPrivilege(undefined)).toBeNull();
  });
});

describe('notFoundSegment / suggestEntitySet (#291)', () => {
  it('reads the segment Dataverse could not resolve', () => {
    expect(notFoundSegment("Resource not found for the segment 'acounts'.")).toBe('acounts');
    expect(notFoundSegment('0x80040217: account With Id = 1 Does Not Exist')).toBeNull();
  });

  it('suggests the closest entity set', () => {
    expect(suggestEntitySet('acounts', ENTITIES)?.entitySetName).toBe('accounts');
    expect(suggestEntitySet('systemuser', ENTITIES)?.entitySetName).toBe('systemusers');
    expect(suggestEntitySet('cso_widgts', ENTITIES)?.entitySetName).toBe('cso_widgets');
  });

  it('matches case and the logical name first', () => {
    expect(suggestEntitySet('Accounts', ENTITIES)?.entitySetName).toBe('accounts');
    expect(suggestEntitySet('contact', ENTITIES)?.entitySetName).toBe('contacts');
  });

  it('offers nothing for an exact entity set, a far guess, or a tie', () => {
    expect(suggestEntitySet('accounts', ENTITIES)).toBeNull();
    expect(suggestEntitySet('invoices', ENTITIES)).toBeNull();
    // 'team' and 'teams'-like sets equally close: no guess.
    expect(suggestEntitySet('xeams', [entity('a', 'teams'), entity('b', 'reams')])).toBeNull();
  });

  it('only offers what it is given (the scoped list)', () => {
    expect(suggestEntitySet('acounts', [entity('contact', 'contacts')])).toBeNull();
  });
});

describe('replacePathSegment (#291)', () => {
  const base = 'https://org.crm.dynamics.com/api/data/v9.2/';
  it('replaces the segment and keeps the query and key', () => {
    expect(
      replacePathSegment(`${base}acounts?$top=5&$filter=name eq 'acounts'`, 'acounts', 'accounts'),
    ).toBe(`${base}accounts?$top=5&$filter=name eq 'acounts'`);
    expect(
      replacePathSegment(
        `${base}acounts(00000000-0000-0000-0000-000000000001)`,
        'acounts',
        'accounts',
      ),
    ).toBe(`${base}accounts(00000000-0000-0000-0000-000000000001)`);
  });

  it('leaves the URL alone when the segment is absent', () => {
    expect(replacePathSegment(`${base}contacts`, 'acounts', 'accounts')).toBe(`${base}contacts`);
  });
});

describe('errorReport (#291)', () => {
  const failed = (over: Partial<ExecResult> = {}): ExecResult => ({
    status: 404,
    statusText: 'Not Found',
    ms: 120,
    bytes: 0,
    ok: false,
    outcome: '404',
    body: {
      error: { code: '0x8006088a', message: "Resource not found for the segment 'acounts'." },
    },
    headers: {
      Authorization: 'Bearer abc.def',
      MSCRMCallerID: '00000000-0000-0000-0000-000000000009',
    },
    ...over,
  });
  const at = new Date('2026-10-30T12:00:00Z');

  it('has the request, status, code and message', () => {
    expect(
      errorReport({
        result: failed(),
        method: 'GET',
        requestUrl: 'https://org/api/data/v9.2/acounts',
        at,
      }),
    ).toBe(
      [
        'Dataverse Request Studio: error report',
        'Time: 2026-10-30T12:00:00.000Z',
        'Request: GET https://org/api/data/v9.2/acounts',
        'Status: 404 Not Found',
        'Code: 0x8006088a',
        "Message: Resource not found for the segment 'acounts'.",
        'Duration: 120 ms',
      ].join('\n'),
    );
  });

  it('never includes headers, and masks a bearer token in the message', () => {
    const report = errorReport({
      result: failed({
        body: { error: { message: 'Token Bearer eyJ0eXAi.abc-def_1 was rejected' } },
      }),
      at,
    });
    expect(report).not.toMatch(
      /eyJ0eXAi|abc\.def|MSCRMCallerID|00000000-0000-0000-0000-000000000009/,
    );
    expect(report).toContain('Message: Token Bearer [removed] was rejected');
  });

  it('names the operation of a write that has no single URL', () => {
    expect(errorReport({ result: failed(), operation: 'update on account', at })).toContain(
      'Time: 2026-10-30T12:00:00.000Z\nOperation: update on account\nStatus: 404 Not Found',
    );
  });

  it('says when the host reported no status', () => {
    expect(errorReport({ result: failed({ status: 0, statusText: '' }), at })).toContain(
      'Status: none (the host reported no HTTP status)',
    );
  });
});
