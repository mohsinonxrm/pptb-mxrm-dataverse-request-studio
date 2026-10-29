// #221: function parameter values in the URL can't add parameters to the call.
import { describe, expect, it } from 'vitest';
import { encodeParamForUrl, invalidFunctionParamReason } from './executeBuilders';
import type { ActionParam } from '../mock/actionsCsdl';

const ID = '00000000-0000-0000-0000-0000000000a1';
const p = (name: string, type: string, extra: Partial<ActionParam> = {}) =>
  ({ name, type, required: false, ...extra }) as ActionParam;

describe('encodeParamForUrl', () => {
  it('writes GUIDs and dates unquoted only when they are one', () => {
    expect(encodeParamForUrl(p('Id', 'Edm.Guid'), ID, false)).toBe(ID);
    expect(encodeParamForUrl(p('Id', 'Edm.Guid'), "x,Other='y'", false)).toBe("'x,Other=''y'''");
    expect(encodeParamForUrl(p('At', 'Edm.DateTimeOffset'), '2024-01-31T10:00:00Z', false)).toBe(
      '2024-01-31T10:00:00Z',
    );
    expect(encodeParamForUrl(p('At', 'Edm.DateTimeOffset'), '2024-01-01),X=(1', false)).toBe(
      "'2024-01-01),X=(1'",
    );
  });

  it("escapes a record reference's id inside its quoted @odata.id", () => {
    expect(
      encodeParamForUrl(
        p('Target', 'EntityReference'),
        { id: "x')/y", entityType: 'zz_unknown' },
        true,
      ),
    ).toBe("{'@odata.id':'zz_unknowns(x'')/y)'}");
  });
});

describe('invalidFunctionParamReason', () => {
  const params = [
    p('Id', 'Edm.Guid'),
    p('Ids', 'Collection(Edm.Guid)'),
    p('At', 'Edm.DateTimeOffset'),
    p('Target', 'EntityReference'),
    p('Name', 'Edm.String'),
  ];

  it('passes valid and empty values', () => {
    expect(
      invalidFunctionParamReason(params, {
        Id: ID,
        Ids: [ID],
        At: '2024-01-31',
        Target: { id: ID, entityType: 'account' },
        Name: "it's fine",
      }),
    ).toBeNull();
    expect(invalidFunctionParamReason(params, {})).toBeNull();
  });

  it('names the first parameter that is not what its type needs', () => {
    expect(invalidFunctionParamReason(params, { Ids: [ID, 'nope'] })).toBe(
      'Parameter Ids: "nope" isn\'t a GUID.',
    );
    expect(invalidFunctionParamReason(params, { At: 'tomorrow' })).toBe(
      'Parameter At: "tomorrow" isn\'t an ISO 8601 date or date-time.',
    );
    expect(invalidFunctionParamReason(params, { Target: { id: 'x', entityType: 'account' } })).toBe(
      'Parameter Target: "x" isn\'t a GUID.',
    );
  });
});
