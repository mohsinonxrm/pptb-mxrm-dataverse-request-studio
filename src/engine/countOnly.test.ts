// Count only (#272): the one-row `$count=true` request, pasting a `/$count`
// URL, and the headers the exports send with it.
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { buildRetrieveMultiple, COUNT_ONLY_PREFER } from './urlBuilder';
import { parseODataUrl } from './odataParser';
import { generateCode } from './codeGenerators';
import { composeHeaders } from './savedToCollection';
import { composePartHeaders } from './batch/savedToBatch';
import { makeExpandSpec } from '../editors/ExpandEditor';
import { emptyApply } from '../editors/ApplyEditor';
import { emptyPrefer } from '../editors/PreferEditor';
import {
  account,
  group,
  registerFixtureTables,
  retrieveMultipleState,
  rule,
} from '../test/fixtures';

vi.mock('../mock/environment', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../mock/environment')>()),
  ENV: {
    name: 'Test',
    host: 'contoso.crm.dynamics.com',
    apiBase: '/api/data/v9.2',
    user: '',
  },
}));

const resolve = (set: string) => (set === account.entitySetName ? account.logicalName : undefined);

beforeEach(() => {
  registerFixtureTables();
});

describe('buildRetrieveMultiple with countOnly', () => {
  it('asks for one id with $count=true and $filter, whatever else is set', () => {
    const b = buildRetrieveMultiple(
      retrieveMultipleState({
        countOnly: true,
        select: ['name'],
        filter: group('and', [rule('numberofemployees', 'gt', '10')]),
        orderby: [{ col: 'name', dir: 'asc' }] as never,
        top: 5,
        expand: [makeExpandSpec('primarycontactid')],
        prefer: { ...emptyPrefer(), maxpagesize: 50 },
      }),
    );
    expect(b.relativeUrl).toBe(
      '/api/data/v9.2/accounts?$select=accountid&$filter=numberofemployees gt 10&$top=1&$count=true',
    );
    expect(b.relativeNoBase).toBe(
      '/accounts?$select=accountid&$filter=numberofemployees gt 10&$top=1&$count=true',
    );
    expect(b.entitySet).toBe('accounts');
  });

  it('leaves out $apply too, and $filter when there is none', () => {
    const b = buildRetrieveMultiple(
      retrieveMultipleState({
        countOnly: true,
        apply: { ...emptyApply(), enabled: true } as never,
      }),
    );
    expect(b.relativeNoBase).toBe('/accounts?$select=accountid&$top=1&$count=true');
  });

  it('leaves the rows request alone when off', () => {
    const b = buildRetrieveMultiple(retrieveMultipleState({ countOnly: false, top: 5 }));
    expect(b.relativeNoBase).toBe('/accounts?$top=5');
  });
});

describe('parseODataUrl with /$count', () => {
  it('turns Count only on for a bare /$count', async () => {
    const r = await parseODataUrl(
      'https://contoso.crm.dynamics.com/api/data/v9.2/accounts/$count',
      resolve,
    );
    expect(r.errors).toEqual([]);
    expect(r.warnings).toEqual([]);
    expect(r.parsed?.table).toBe('account');
    expect(r.parsed?.countOnly).toBe(true);
  });

  it('keeps a $filter and says how it is sent', async () => {
    const r = await parseODataUrl('accounts/$count?$filter=numberofemployees gt 10', resolve);
    expect(r.errors).toEqual([]);
    expect(r.parsed?.countOnly).toBe(true);
    expect(r.parsed?.filter.rules).toHaveLength(1);
    expect(r.warnings.map((w) => w.message).join(' ')).toMatch(
      /doesn't accept \$filter on \/\$count/,
    );
    const b = buildRetrieveMultiple(retrieveMultipleState({ ...r.parsed!, dirty: new Set() }));
    expect(b.relativeNoBase).toBe(
      '/accounts?$select=accountid&$filter=numberofemployees gt 10&$top=1&$count=true',
    );
  });

  it.each([
    'accounts(00000000-0000-0000-0000-0000000000a1)/contact_customer_accounts/$count',
    'accounts(00000000-0000-0000-0000-0000000000a1)/$count',
  ])('refuses %s rather than count the whole table', async (url) => {
    const r = await parseODataUrl(url, resolve);
    expect(r.ok).toBe(false);
    expect(r.errors.map((e) => e.message).join(' ')).toMatch(/counts a table's rows/);
  });

  it('leaves countOnly off for a rows URL', async () => {
    const r = await parseODataUrl('accounts?$count=true', resolve);
    expect(r.parsed?.countOnly).toBeUndefined();
    expect(r.parsed?.countOn).toBe(true);
  });
});

describe('Count only in the Code tab and exports', () => {
  it('every format sends the request that runs', () => {
    const built = buildRetrieveMultiple(retrieveMultipleState({ countOnly: true }));
    const i = { method: 'GET', built, headers: { Prefer: COUNT_ONLY_PREFER } };
    for (const fmt of ['fetch', 'curl', 'powershell', 'csharp', 'python'] as const) {
      const code = generateCode(fmt, i);
      expect(code, fmt).toContain('$top=1');
      expect(code, fmt).toContain('$count=true');
      expect(code, fmt).not.toContain('/$count');
    }
  });

  it('collections and $batch ask for the limit annotation, never a page size', () => {
    const s = retrieveMultipleState({
      countOnly: true,
      prefer: { ...emptyPrefer(), maxpagesize: 50 },
    });
    expect(composeHeaders(s).Prefer).toBe(COUNT_ONLY_PREFER);
    expect(composePartHeaders(s).Prefer).toBe(COUNT_ONLY_PREFER);
    const rows = retrieveMultipleState({ prefer: { ...emptyPrefer(), maxpagesize: 50 } });
    expect(composeHeaders(rows).Prefer).toMatch(/odata\.maxpagesize=50/);
  });
});
