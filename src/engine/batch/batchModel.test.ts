import { describe, expect, it } from 'vitest';
import {
  compileBatch,
  deserializeBatch,
  itemToWire,
  newGroup,
  newItem,
  normalizePath,
  planDispatch,
  serializeBatch,
} from './batchModel';

const ORG = 'https://org.crm.dynamics.com';

describe('itemToWire', () => {
  it('builds each operation type', () => {
    expect(
      itemToWire(newItem('create', { entitySet: 'accounts', body: '{"name":"A"}' })).wire,
    ).toEqual({ method: 'POST', path: 'accounts', headers: {}, body: '{\r\n  "name": "A"\r\n}' });
    expect(
      itemToWire(newItem('update', { entitySet: 'accounts', recordId: '{ABC}', body: '{}' })).wire
        ?.path,
    ).toBe('accounts(ABC)');
    expect(itemToWire(newItem('delete', { entitySet: 'accounts', recordId: '1' })).wire).toEqual({
      method: 'DELETE',
      path: 'accounts(1)',
      headers: {},
      body: undefined,
    });
    expect(
      itemToWire(newItem('retrieve', { entitySet: 'accounts', query: '?$select=name&$top=3' })).wire
        ?.path,
    ).toBe('accounts?$select=name&$top=3');
    const assoc = itemToWire(
      newItem('associate', {
        entitySet: 'accounts',
        recordId: '1',
        navProperty: 'contact_customer_accounts',
        targetSet: 'contacts',
        targetId: '2',
      }),
      ORG,
    ).wire;
    expect(assoc?.path).toBe('accounts(1)/contact_customer_accounts/$ref');
    expect(JSON.parse(assoc!.body!)).toEqual({
      '@odata.id': `${ORG}/api/data/v9.2/contacts(2)`,
    });
    expect(
      itemToWire(newItem('custom', { method: 'GET', url: '/api/data/v9.2/WhoAmI' })).wire,
    ).toMatchObject({ method: 'GET', path: 'WhoAmI' });
  });

  it('uses $n references as the whole record segment', () => {
    expect(itemToWire(newItem('update', { recordId: '$1', body: '{}' })).wire?.path).toBe('$1');
  });

  it('reports invalid input', () => {
    const r = itemToWire(newItem('create', { entitySet: 'accounts', body: '{nope' }));
    expect(r.wire).toBeNull();
    expect(r.problems[0].message).toMatch(/valid JSON/);
    expect(
      itemToWire(newItem('update', { entitySet: 'a', recordId: '1', body: '[1]' })).problems[0]
        .message,
    ).toMatch(/JSON object/);
    expect(
      itemToWire(newItem('custom', { method: 'POST', url: '$batch' })).problems[0].message,
    ).toMatch(/another \$batch/);
    expect(
      itemToWire(
        newItem('delete', {
          entitySet: 'a',
          recordId: '1',
          headers: [{ name: 'Bad Name', value: 'x', enabled: true }],
        }),
      ).problems[0].message,
    ).toMatch(/Invalid header name/);
  });

  it('keeps enabled headers only', () => {
    const w = itemToWire(
      newItem('delete', {
        entitySet: 'a',
        recordId: '1',
        headers: [
          { name: 'If-Match', value: '*', enabled: true },
          { name: 'MSCRM.BypassBusinessLogicExecution', value: 'CustomSync', enabled: false },
        ],
      }),
    ).wire;
    expect(w?.headers).toEqual({ 'If-Match': '*' });
  });
});

describe('normalizePath', () => {
  it('strips the host and API base', () => {
    expect(normalizePath('https://x.crm.dynamics.com/api/data/v9.2/accounts?$top=1')).toBe(
      'accounts?$top=1',
    );
    expect(normalizePath('/accounts')).toBe('accounts');
  });
});

describe('compileBatch', () => {
  it('numbers Content-IDs across groups and flags GET in a change set', () => {
    const groups = [
      newGroup(true, [
        newItem('create', { entitySet: 'accounts', body: '{"name":"A"}' }),
        newItem('retrieve', { entitySet: 'accounts' }),
      ]),
      newGroup(false, [newItem('retrieve', { entitySet: 'contacts' })]),
    ];
    const { items, issues } = compileBatch(groups);
    expect(items.map((i) => i.contentId)).toEqual(['1', '2', '3']);
    expect(issues).toHaveLength(1);
    expect(issues[0]).toMatchObject({ severity: 'error', target: groups[0].items[1].id });
  });

  it('allows body references to earlier requests in the same change set only', () => {
    const ok = compileBatch([
      newGroup(true, [
        newItem('create', { entitySet: 'contacts', body: '{"firstname":"a"}' }),
        newItem('create', {
          entitySet: 'accounts',
          body: '{"name":"b","primarycontactid@odata.bind":"$1"}',
        }),
      ]),
    ]);
    expect(ok.issues).toEqual([]);

    const forward = compileBatch([
      newGroup(true, [
        newItem('create', { entitySet: 'accounts', body: '{"x@odata.bind":"$2"}' }),
        newItem('create', { entitySet: 'contacts', body: '{}' }),
      ]),
    ]);
    expect(forward.issues[0].message).toMatch(/earlier request in the same change set/);

    const independent = compileBatch([
      newGroup(false, [
        newItem('create', { entitySet: 'contacts', body: '{}' }),
        newItem('create', { entitySet: 'accounts', body: '{"x@odata.bind":"$1"}' }),
      ]),
    ]);
    expect(independent.issues[0].message).toMatch(/only be referenced inside a change set/);
  });

  it('warns about URL references the host prefixes', () => {
    const { issues } = compileBatch([
      newGroup(true, [
        newItem('create', { entitySet: 'contacts', body: '{}' }),
        newItem('update', { recordId: '$1', body: '{"lastname":"x"}' }),
      ]),
    ]);
    expect(issues).toHaveLength(1);
    expect(issues[0].severity).toBe('warning');
  });

  it('caps the batch at 1000 requests', () => {
    const items = Array.from({ length: 1001 }, () => newItem('retrieve', { entitySet: 'a' }));
    const { issues } = compileBatch([newGroup(false, items)]);
    expect(issues.some((i) => /at most 1000/.test(i.message))).toBe(true);
  });

  it('needs at least one request', () => {
    expect(compileBatch([newGroup(true)]).issues[0].message).toBe('Add a request.');
  });
});

describe('planDispatch', () => {
  it('sends each change set on its own and merges consecutive independent groups', () => {
    const groups = [
      newGroup(true, [newItem('delete', { entitySet: 'a', recordId: '1' })]),
      newGroup(true, [newItem('delete', { entitySet: 'a', recordId: '2' })]),
      newGroup(false, [newItem('retrieve', { entitySet: 'a' })]),
      newGroup(false, [newItem('retrieve', { entitySet: 'b' })]),
      newGroup(true, []),
      newGroup(true, [
        newItem('delete', { entitySet: 'a', recordId: '3' }),
        newItem('delete', { entitySet: 'a', recordId: '4' }),
      ]),
    ];
    const plan = planDispatch(compileBatch(groups).items);
    expect(plan.map((s) => [s.transaction, s.items.map((i) => i.contentId)])).toEqual([
      [true, ['1']],
      [true, ['2']],
      [false, ['3', '4']],
      [true, ['5', '6']],
    ]);
  });
});

describe('serialize / deserialize', () => {
  it('drops dirty and backfills missing item fields', () => {
    const state = {
      groups: [newGroup(true, [newItem('create', { entitySet: 'accounts' })])],
      continueOnError: true,
      dirty: new Set(['x']),
    };
    const snap = JSON.parse(JSON.stringify(serializeBatch(state)));
    expect(snap.dirty).toBeUndefined();
    delete snap.groups[0].items[0].headers;
    const back = deserializeBatch(snap);
    expect(back.continueOnError).toBe(true);
    expect(back.groups[0].items[0].headers).toEqual([]);
    expect(back.dirty.size).toBe(0);
  });
});
