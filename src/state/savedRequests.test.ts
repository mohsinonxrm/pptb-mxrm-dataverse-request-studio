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
