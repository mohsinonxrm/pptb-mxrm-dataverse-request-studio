// Saved-request storage: reading never loses data, and every mode is known (#220).
import { beforeEach, describe, expect, it } from 'vitest';
import {
  SAVED_BACKUP_KEY_PREFIX,
  SAVED_MODE_IDS,
  SAVED_REQUESTS_KEY_PREFIX,
  backupUnreadable,
  buildSavedExport,
  getSavedRequestsKey,
  listOtherOrgLibraries,
  parseSavedImport,
  persistSaved,
  readLibrary,
  type SavedRequest,
} from './savedRequests';

const entry = (id: string, modeId: string, extra: Partial<SavedRequest> = {}): SavedRequest =>
  ({ id, name: id, modeId, orgScope: 'default', state: {}, savedAt: 1, ...extra }) as SavedRequest;

beforeEach(() => localStorage.clear());

describe('SAVED_MODE_IDS', () => {
  it('includes the 1.2.0 modes', () => {
    for (const m of ['batch', 'execute-multiple', 'bulk-write', 'bulk-delete'])
      expect(SAVED_MODE_IDS.has(m)).toBe(true);
  });
});

describe('readLibrary', () => {
  const key = `${SAVED_REQUESTS_KEY_PREFIX}:org.example`;

  it('reads a missing bucket as an empty library', () => {
    expect(readLibrary(key, 'org.example')).toEqual({
      entries: [],
      preserved: [],
      unreadable: null,
    });
  });

  it('reports text that is not JSON, or not an array, as unreadable', () => {
    localStorage.setItem(key, '[{"id":"a",');
    expect(readLibrary(key, 'org.example')).toMatchObject({
      entries: [],
      unreadable: '[{"id":"a",',
    });
    localStorage.setItem(key, '{"id":"a"}');
    expect(readLibrary(key, 'org.example').unreadable).toBe('{"id":"a"}');
  });

  it('keeps items it cannot use as preserved raw values', () => {
    const items = [
      entry('ok', 'retrieve-multiple'),
      entry('batch', 'batch'),
      entry('future-mode', 'graphql-query'),
      entry('newer', 'create', { schemaVersion: 2 }),
      { id: 'broken' },
      'junk',
    ];
    localStorage.setItem(key, JSON.stringify(items));
    const lib = readLibrary(key, 'org.example');
    expect(lib.unreadable).toBeNull();
    expect(lib.entries.map((e) => e.id)).toEqual(['ok', 'batch']);
    expect(lib.entries[0]).toMatchObject({ orgScope: 'default', schemaVersion: 1 });
    expect(lib.preserved).toEqual(items.slice(2));
  });
});

describe('persistSaved', () => {
  it('writes the entries followed by the preserved items, unchanged', () => {
    const preserved = [entry('newer', 'create', { schemaVersion: 2 }), { id: 'broken' }];
    persistSaved([entry('ok', 'update')], preserved);
    expect(JSON.parse(localStorage.getItem(getSavedRequestsKey())!)).toEqual([
      entry('ok', 'update'),
      ...preserved,
    ]);
  });
});

describe('backupUnreadable', () => {
  it('copies the text once, outside the library prefix', () => {
    const k1 = backupUnreadable('org.example', 'not json');
    const k2 = backupUnreadable('org.example', 'not json');
    expect(k1).toBe(k2);
    expect(k1!.startsWith(`${SAVED_BACKUP_KEY_PREFIX}:org.example:`)).toBe(true);
    expect(localStorage.getItem(k1!)).toBe('not json');
    expect(k1!.startsWith(`${SAVED_REQUESTS_KEY_PREFIX}:`)).toBe(false);
  });
});

describe('import and the other-connections view', () => {
  it('import keeps $batch, ExecuteMultiple and Bulk entries', () => {
    const file = buildSavedExport(
      [
        entry('b', 'batch'),
        entry('em', 'execute-multiple'),
        entry('bw', 'bulk-write'),
        entry('bd', 'bulk-delete'),
      ],
      'org.example',
    );
    const r = parseSavedImport(file);
    expect(r.invalid).toBe(0);
    expect(r.entries.map((e) => e.id)).toEqual(['b', 'em', 'bw', 'bd']);
  });

  it('lists other connections’ $batch entries and never a backup key', () => {
    localStorage.setItem(
      `${SAVED_REQUESTS_KEY_PREFIX}:other.example`,
      JSON.stringify([entry('b', 'batch')]),
    );
    backupUnreadable('third.example', 'not json');
    const libs = listOtherOrgLibraries();
    expect(libs.map((l) => l.scope)).toEqual(['other.example']);
    expect(libs[0].entries.map((e) => e.id)).toEqual(['b']);
  });
});

describe('import checks ids and names (#221)', () => {
  const ID = '00000000-0000-0000-0000-0000000000a1';
  const withState = (id: string, modeId: string, state: Record<string, unknown>) =>
    entry(id, modeId, { state });

  it('skips entries whose ids or names would change the URL, and says why', () => {
    const file = buildSavedExport(
      [
        withState('ok', 'retrieve-single', { table: 'account', recordId: ID }),
        withState('path', 'retrieve-single', { table: 'account', recordId: `${ID})/x` }),
        withState('nav', 'associate', {
          table: 'account',
          sourceId: ID,
          navProperty: 'nav?$top=1',
          targets: [ID],
        }),
        withState('target', 'disassociate', {
          table: 'account',
          sourceId: ID,
          navProperty: 'contact_customer_accounts',
          targetIds: [ID, 'bad'],
        }),
        withState('query', 'predefined-query', { table: 'account', queryId: 'x&$top=1' }),
        withState('table', 'update', { table: 'accounts(x)', recordId: ID }),
        withState('type', 'merge', { table: 'account', targetId: 42 }),
      ],
      'org.example',
    );
    const r = parseSavedImport(file);
    expect(r.entries.map((e) => e.id)).toEqual(['ok']);
    expect(r.invalid).toBe(6);
    expect(r.problems).toEqual([
      `"path": Record id "${ID})/x" isn't a GUID (00000000-0000-0000-0000-000000000000).`,
      '"nav": Navigation property "nav?$top=1" isn\'t a valid name (letters, digits and underscores).',
      '"target": Target record id "bad" isn\'t a GUID (00000000-0000-0000-0000-000000000000).',
      '"query": Query id "x&$top=1" isn\'t a GUID (00000000-0000-0000-0000-000000000000).',
      '"table": Table "accounts(x)" isn\'t a valid name (letters, digits and underscores).',
      '"type": Master record id isn\'t text.',
    ]);
  });

  it("keeps empty ids, an upsert by alternate key, and batch items' alternate keys", () => {
    const file = buildSavedExport(
      [
        withState('empty', 'delete', { table: 'account', recordId: null }),
        withState('alt', 'upsert', {
          table: 'account',
          key: { kind: 'alternate', values: { accountnumber: "A'1" } },
        }),
        withState('batch', 'batch', { groups: [{ items: [{ recordId: "accountnumber='A1'" }] }] }),
      ],
      'org.example',
    );
    const r = parseSavedImport(file);
    expect(r.entries.map((e) => e.id)).toEqual(['empty', 'alt', 'batch']);
    expect(r.problems).toEqual([]);
  });
});

describe('import checks query names (#221 review)', () => {
  it('skips a Retrieve Multiple whose filter column, $select or $expand is not a name', () => {
    const rule = (col: string) => ({
      id: 'r',
      type: 'rule',
      col,
      op: 'eq',
      val: 'x',
    });
    const group = (col: string) => ({
      id: 'g',
      type: 'group',
      combinator: 'and',
      rules: [rule(col)],
    });
    const base = {
      table: 'account',
      select: ['name'],
      orderby: [],
      expand: [],
      filter: group('name'),
    };
    const file = buildSavedExport(
      [
        entry('ok', 'retrieve-multiple', { state: base }),
        entry('col', 'retrieve-multiple', {
          state: { ...base, filter: group("name eq 'a' or name") },
        }),
        entry('sel', 'retrieve-multiple', { state: { ...base, select: ['name,revenue'] } }),
        entry('exp', 'retrieve-multiple', { state: { ...base, expand: [{ nav: 'x?$top=1' }] } }),
        entry('bad', 'retrieve-multiple', { state: { ...base, filter: { rules: [null] } } }),
      ],
      'org.example',
    );
    const r = parseSavedImport(file);
    expect(r.entries.map((e) => e.id)).toEqual(['ok']);
    expect(r.problems.map((p) => p.split(':')[0])).toEqual(['"col"', '"sel"', '"exp"', '"bad"']);
  });
});
