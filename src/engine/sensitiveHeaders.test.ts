// Credential and impersonation headers are never saved or exported as typed (#217).
import { beforeEach, describe, expect, it, vi } from 'vitest';
import {
  impersonationVar,
  isCredentialHeader,
  stripCredentialRows,
  withoutCredentialHeaders,
} from './sensitiveHeaders';
import { buildCollection, type CollectionInput } from './collectionFormats';
import { generateCode, type CodegenInputs } from './codeGenerators';
import { toWireHeaders } from './dataverseExecutor';
import { itemToWire, newItem } from './batch/batchModel';
import { buildSavedExport, type SavedRequest } from '../state/savedRequests';

vi.mock('../env/environment', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../env/environment')>()),
  ENV: { name: 'Test', host: 'contoso.crm.dynamics.com', apiBase: '/api/data/v9.2', user: '' },
}));

const CALLER = '11111111-2222-3333-4444-555555555555';
const TOKEN = 'Bearer eyJ0eXAiOiJKV1QiLCJhbGciOi.secret';
const row = (name: string, value: string, enabled = true) => ({ id: name, name, value, enabled });

describe('header classification', () => {
  it('treats Authorization, Proxy-Authorization and Cookie as credentials, in any case', () => {
    for (const n of [
      'Authorization',
      'authorization',
      ' AUTHORIZATION ',
      'Proxy-Authorization',
      'Cookie',
    ])
      expect(isCredentialHeader(n)).toBe(true);
    for (const n of ['MSCRMCallerID', 'If-Match', 'Prefer', 'X-Authorization-Info'])
      expect(isCredentialHeader(n)).toBe(false);
  });

  it('maps impersonation headers to collection variables', () => {
    expect(impersonationVar('MSCRMCallerID')).toBe('callerId');
    expect(impersonationVar('callerobjectid')).toBe('callerObjectId');
    expect(impersonationVar('Prefer')).toBeUndefined();
  });

  it('drops credential headers from a header map', () => {
    expect(
      withoutCredentialHeaders({ Authorization: TOKEN, Cookie: 'a=b', 'If-Match': '*' }),
    ).toEqual({ 'If-Match': '*' });
  });
});

describe('stripCredentialRows', () => {
  it('removes credential rows from every header-row array in a saved state', () => {
    const state = {
      table: 'account',
      headers: [
        row('Authorization', TOKEN),
        row('MSCRMCallerID', CALLER),
        row('Cookie', 'x', false),
      ],
      groups: [
        { items: [{ kind: 'custom', headers: [row('authorization', TOKEN), row('Prefer', 'x')] }] },
      ],
    };
    expect(stripCredentialRows(state)).toEqual({
      table: 'account',
      headers: [row('MSCRMCallerID', CALLER)],
      groups: [{ items: [{ kind: 'custom', headers: [row('Prefer', 'x')] }] }],
    });
  });

  it('returns the same object when there is nothing to remove', () => {
    const state = { headers: [row('Prefer', 'x')], fieldValues: { name: 'A' } };
    expect(stripCredentialRows(state)).toBe(state);
  });

  it('also cleans the saved-requests export file', () => {
    const e = {
      id: 'a',
      name: 'a',
      modeId: 'retrieve-multiple',
      orgScope: 'x',
      state: { headers: [row('Authorization', TOKEN)] },
      savedAt: 1,
    } as SavedRequest;
    expect(buildSavedExport([e], 'x')).not.toContain('Authorization');
  });
});

describe('collection exports', () => {
  const input: CollectionInput = {
    name: 'DRS',
    baseUrl: 'https://contoso.crm.dynamics.com',
    items: [
      {
        kind: 'request',
        name: 'Accounts',
        method: 'GET',
        path: '/api/data/v9.2/accounts',
        headers: {
          Authorization: TOKEN,
          Cookie: 'a=b',
          MSCRMCallerID: CALLER,
          Prefer: 'odata.maxpagesize=5',
        },
      },
    ],
  };

  it.each(['postman', 'insomnia', 'bruno'] as const)(
    '%s: no credential, the caller id becomes a declared variable',
    (fmt) => {
      const text = buildCollection(fmt, input);
      expect(text).not.toContain(TOKEN);
      expect(text).not.toContain('a=b');
      expect(text).not.toContain(CALLER);
      expect(text).toContain('odata.maxpagesize=5');
      const json = JSON.parse(text);
      if (fmt === 'postman') {
        expect(json.item[0].request.header).toContainEqual({
          key: 'MSCRMCallerID',
          value: '{{callerId}}',
          type: 'text',
        });
        expect(json.variable).toContainEqual({ key: 'callerId', value: '', type: 'string' });
      } else if (fmt === 'insomnia') {
        const req = json.resources.find((r: { _type: string }) => r._type === 'request');
        expect(req.headers).toContainEqual({ name: 'MSCRMCallerID', value: '{{ _.callerId }}' });
        const env = json.resources.find((r: { _type: string }) => r._type === 'environment');
        expect(env.data.callerId).toBe('');
      } else {
        expect(json.items[0].request.headers).toContainEqual({
          name: 'MSCRMCallerID',
          value: '{{callerId}}',
          enabled: true,
        });
        expect(json.environments[0].variables.map((v: { name: string }) => v.name)).toContain(
          'callerId',
        );
      }
    },
  );
});

describe('Code tab and $batch parts', () => {
  const codegen: CodegenInputs = {
    method: 'GET',
    built: {
      relativeUrl: '/api/data/v9.2/accounts',
      relativeNoBase: '/accounts',
      bytes: 0,
      queryParts: [],
      entitySet: 'accounts',
      entityLogical: 'account',
    },
    headers: { Authorization: TOKEN, Cookie: 'a=b', 'OData-Version': '4.0' },
  };

  it.each(['fetch', 'curl', 'powershell', 'csharp', 'python', 'json', 'xhr'] as const)(
    '%s never carries the typed credential',
    (fmt) => {
      const code = generateCode(fmt, codegen);
      expect(code).not.toContain(TOKEN);
      expect(code).not.toContain('a=b');
    },
  );

  let warnings: string[];
  beforeEach(() => {
    const r = itemToWire(
      newItem('custom', {
        method: 'GET',
        url: '/api/data/v9.2/WhoAmI',
        headers: [{ name: 'Authorization', value: TOKEN, enabled: true }],
      }),
    );
    expect(r.wire?.headers).toEqual({});
    warnings = r.problems.filter((p) => p.severity === 'warning').map((p) => p.message);
  });

  it("doesn't send a credential on a $batch part, and says so", () => {
    expect(warnings).toEqual([
      "Authorization isn't sent: Power Platform ToolBox authenticates the batch.",
    ]);
  });
});

describe('requests and OpenAPI', () => {
  it('never hands a credential header to the host when a request runs', () => {
    expect(
      toWireHeaders({
        Cookie: 'a=b',
        'Proxy-Authorization': 'x',
        Authorization: TOKEN,
        'If-Match': '*',
      }),
    ).toEqual({ 'If-Match': '*' });
    expect(toWireHeaders({ Cookie: 'a=b' })).toBeUndefined();
  });

  it("OpenAPI doesn't write the caller id as an example", () => {
    const code = generateCode('openapi', {
      method: 'GET',
      built: {
        relativeUrl: '/api/data/v9.2/accounts',
        relativeNoBase: '/accounts',
        bytes: 0,
        queryParts: [],
        entitySet: 'accounts',
        entityLogical: 'account',
      },
      headers: { MSCRMCallerID: CALLER, Cookie: 'a=b' },
    });
    expect(code).toContain('MSCRMCallerID');
    expect(code).not.toContain(CALLER);
    expect(code).not.toContain('a=b');
  });
});
