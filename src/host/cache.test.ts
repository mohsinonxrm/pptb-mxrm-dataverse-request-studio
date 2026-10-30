// The metadata cache and the loaders over it (#235): the one-hour TTL,
// in-flight de-duplication, failures not cached, and one cache per org (#52).
// Empty solution sets select nothing (#44).
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { AttributeMetadata, EntityMetadata, SolutionComponent } from './pptbClient';

const org = vi.hoisted(() => ({ current: 'org-a' }));
vi.mock('./activeOrg', () => ({
  activeOrg: () => org.current,
  onActiveOrgChange: () => () => {},
}));

const client = vi.hoisted(() => ({
  getAllEntities: vi.fn(),
  getEntityAttributes: vi.fn(),
  getSolutionComponents: vi.fn(),
}));
vi.mock('./pptbClient', async (importOriginal) => ({
  ...(await importOriginal<typeof import('./pptbClient')>()),
  ...client,
}));

import { metadataCache, pinMetadataCache } from './cache';
import {
  loadAllEntities,
  loadEntityAttributes,
  loadSolutionEntityNames,
} from './dataverseMetadata';

const HOUR = 60 * 60 * 1000;

const entity = (LogicalName: string): EntityMetadata => ({
  LogicalName,
  SchemaName: LogicalName,
  EntitySetName: `${LogicalName}s`,
  PrimaryIdAttribute: `${LogicalName}id`,
  PrimaryNameAttribute: 'name',
  MetadataId: LogicalName,
  ObjectTypeCode: 1,
});

const attribute = (LogicalName: string): AttributeMetadata => ({
  LogicalName,
  SchemaName: LogicalName,
  AttributeType: 'String',
  MetadataId: LogicalName,
});

const component = (msdyn_name: string): SolutionComponent => ({
  msdyn_name,
  msdyn_solutionid: 's1',
  msdyn_componenttype: 1,
});

/** A promise the test settles by hand. */
function deferred<T>() {
  let resolve!: (v: T) => void;
  let reject!: (e: unknown) => void;
  const promise = new Promise<T>((res, rej) => {
    resolve = res;
    reject = rej;
  });
  return { promise, resolve, reject };
}

let orgSeq = 0;
beforeEach(() => {
  // A fresh org per test, so no cache state leaks between tests.
  org.current = `org-${++orgSeq}`;
  vi.resetAllMocks();
  vi.spyOn(console, 'error').mockImplementation(() => {});
});

afterEach(() => {
  vi.useRealTimers();
  vi.restoreAllMocks();
});

describe('TTL', () => {
  beforeEach(() => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date('2026-01-01T09:00:00Z'));
  });

  it('serves an entry for an hour after it was written', () => {
    metadataCache.setEntityMetadata('account', entity('account'));
    vi.advanceTimersByTime(HOUR);
    expect(metadataCache.getEntityMetadata('account')?.LogicalName).toBe('account');
  });

  it('treats an entry older than an hour as missing', () => {
    metadataCache.setEntityMetadata('account', entity('account'));
    vi.advanceTimersByTime(HOUR + 1);
    expect(metadataCache.getEntityMetadata('account')).toBeUndefined();
  });

  it('applies the TTL to every kind of entry', () => {
    metadataCache.setAllEntities([entity('account')]);
    metadataCache.setEntityAttributes('account_false', [attribute('name')]);
    metadataCache.setEntityRelationships('account_false', {
      oneToMany: [],
      manyToOne: [],
      manyToMany: [],
    });
    metadataCache.setEntityKeys('account', []);
    metadataCache.setPublishers([]);
    metadataCache.setPublishersWithSolutions([]);
    metadataCache.setSolutions(['p1'], []);
    metadataCache.setFilteredEntities(['s1'], []);
    metadataCache.setSolutionComponents(['s1'], []);
    vi.advanceTimersByTime(HOUR + 1);
    expect(metadataCache.getAllEntities()).toBeUndefined();
    expect(metadataCache.getEntityAttributes('account_false')).toBeUndefined();
    expect(metadataCache.getEntityRelationships('account_false')).toBeUndefined();
    expect(metadataCache.getEntityKeys('account')).toBeUndefined();
    expect(metadataCache.getPublishers()).toBeUndefined();
    expect(metadataCache.getPublishersWithSolutions()).toBeUndefined();
    expect(metadataCache.getSolutions(['p1'])).toBeUndefined();
    expect(metadataCache.getFilteredEntities(['s1'])).toBeUndefined();
    expect(metadataCache.getSolutionComponents(['s1'])).toBeUndefined();
  });

  it('re-fetches through the loader once the entry has expired', async () => {
    client.getEntityAttributes.mockResolvedValue([attribute('name')]);
    await loadEntityAttributes('account');
    vi.advanceTimersByTime(HOUR - 1);
    await loadEntityAttributes('account');
    expect(client.getEntityAttributes).toHaveBeenCalledTimes(1);
    vi.advanceTimersByTime(2);
    await loadEntityAttributes('account');
    expect(client.getEntityAttributes).toHaveBeenCalledTimes(2);
  });
});

describe('keys', () => {
  it('finds solutions and filtered entities whatever order the ids come in', () => {
    metadataCache.setSolutions(['p2', 'p1'], []);
    metadataCache.setFilteredEntities(['s2', 's1'], [entity('account')]);
    metadataCache.setSolutionComponents(['s2', 's1'], [component('account')]);
    expect(metadataCache.getSolutions(['p1', 'p2'])).toEqual([]);
    expect(metadataCache.getFilteredEntities(['s1', 's2'])).toHaveLength(1);
    expect(metadataCache.getSolutionComponents(['s1', 's2'])).toHaveLength(1);
    expect(metadataCache.getSolutions(['p1'])).toBeUndefined();
  });

  it('keeps an empty publisher list (all solutions) apart from a specific one', () => {
    metadataCache.setSolutions([], []);
    expect(metadataCache.getSolutions([])).toEqual([]);
    expect(metadataCache.getSolutions(['p1'])).toBeUndefined();
  });
});

describe('filterCachedEntitiesByNames', () => {
  it('returns the cached entities with the given names', () => {
    metadataCache.setAllEntities([entity('account'), entity('contact'), entity('lead')]);
    const names = metadataCache.filterCachedEntitiesByNames(['lead', 'account', 'missing']);
    expect(names.map((e) => e.LogicalName)).toEqual(['account', 'lead']);
  });

  it('selects nothing for an empty name list (#44)', () => {
    metadataCache.setAllEntities([entity('account')]);
    expect(metadataCache.filterCachedEntitiesByNames([])).toEqual([]);
  });

  it('selects nothing when the entity list is not loaded', () => {
    expect(metadataCache.filterCachedEntitiesByNames(['account'])).toEqual([]);
  });
});

describe('invalidation', () => {
  it('clear() drops every entry', () => {
    metadataCache.setAllEntities([entity('account')]);
    metadataCache.setEntityMetadata('account', entity('account'));
    metadataCache.setEntityAttributes('account_false', [attribute('name')]);
    metadataCache.setEntityKeys('account', []);
    metadataCache.setPublishers([]);
    metadataCache.setSolutions([], []);
    metadataCache.clear();
    expect(metadataCache.getAllEntities()).toBeUndefined();
    expect(metadataCache.getEntityMetadata('account')).toBeUndefined();
    expect(metadataCache.getEntityAttributes('account_false')).toBeUndefined();
    expect(metadataCache.getEntityKeys('account')).toBeUndefined();
    expect(metadataCache.getPublishers()).toBeUndefined();
    expect(metadataCache.getSolutions([])).toBeUndefined();
  });

  it('clearEntity() drops one table and its attribute, relationship and key entries', () => {
    const rels = { oneToMany: [], manyToOne: [], manyToMany: [] };
    metadataCache.setEntityMetadata('account', entity('account'));
    metadataCache.setEntityAttributes('account_false', [attribute('name')]);
    metadataCache.setEntityAttributes('account_industrycode_optionset', [attribute('x')]);
    metadataCache.setEntityRelationships('account_false', rels);
    metadataCache.setEntityKeys('account', []);
    metadataCache.setEntityMetadata('contact', entity('contact'));
    metadataCache.setEntityAttributes('contact_false', [attribute('fullname')]);
    metadataCache.setAllEntities([entity('account')]);

    metadataCache.clearEntity('account');

    expect(metadataCache.getEntityMetadata('account')).toBeUndefined();
    expect(metadataCache.getEntityAttributes('account_false')).toBeUndefined();
    expect(metadataCache.getEntityAttributes('account_industrycode_optionset')).toBeUndefined();
    expect(metadataCache.getEntityRelationships('account_false')).toBeUndefined();
    expect(metadataCache.getEntityKeys('account')).toBeUndefined();
    expect(metadataCache.getEntityMetadata('contact')).toBeDefined();
    expect(metadataCache.getEntityAttributes('contact_false')).toBeDefined();
    expect(metadataCache.getAllEntities()).toHaveLength(1);
  });
});

describe('one cache per org (#52)', () => {
  it("never serves another org's entries, and keeps them for switching back", () => {
    metadataCache.setEntityMetadata('account', entity('account'));
    const first = org.current;
    org.current = `${first}-other`;
    expect(metadataCache.getEntityMetadata('account')).toBeUndefined();
    org.current = first;
    expect(metadataCache.getEntityMetadata('account')?.LogicalName).toBe('account');
  });

  it('pinMetadataCache() returns the same cache for the same org', () => {
    const a = pinMetadataCache();
    expect(pinMetadataCache()).toBe(a);
    org.current = `${org.current}-other`;
    expect(pinMetadataCache()).not.toBe(a);
  });

  it('a pinned cache keeps writing to its own org after a switch', () => {
    const pinned = pinMetadataCache();
    const first = org.current;
    org.current = `${first}-other`;
    pinned.setEntityMetadata('account', entity('account'));
    expect(metadataCache.getEntityMetadata('account')).toBeUndefined();
    org.current = first;
    expect(metadataCache.getEntityMetadata('account')).toBeDefined();
  });

  it('a load that lands after a switch is cached under the org it started in', async () => {
    const reply = deferred<AttributeMetadata[]>();
    client.getEntityAttributes.mockReturnValueOnce(reply.promise);
    const first = org.current;
    const load = loadEntityAttributes('account');
    org.current = `${first}-other`;
    reply.resolve([attribute('name')]);
    await load;
    expect(metadataCache.getEntityAttributes('account_false')).toBeUndefined();
    org.current = first;
    client.getEntityAttributes.mockClear();
    expect(await loadEntityAttributes('account')).toEqual([attribute('name')]);
    expect(client.getEntityAttributes).not.toHaveBeenCalled();
  });
});

describe('loaders: in-flight de-duplication', () => {
  it('two concurrent attribute loads share one fetch', async () => {
    const reply = deferred<AttributeMetadata[]>();
    client.getEntityAttributes.mockReturnValueOnce(reply.promise);
    const a = loadEntityAttributes('account');
    const b = loadEntityAttributes('account');
    reply.resolve([attribute('name')]);
    expect(await a).toEqual([attribute('name')]);
    expect(await b).toEqual([attribute('name')]);
    expect(client.getEntityAttributes).toHaveBeenCalledTimes(1);
    expect(client.getEntityAttributes).toHaveBeenCalledWith('account', false);
  });

  it('two concurrent entity-list loads share one fetch', async () => {
    const reply = deferred<EntityMetadata[]>();
    client.getAllEntities.mockReturnValueOnce(reply.promise);
    const a = loadAllEntities();
    const b = loadAllEntities();
    reply.resolve([entity('account')]);
    expect(await a).toBe(await b);
    expect(client.getAllEntities).toHaveBeenCalledTimes(1);
    expect(client.getAllEntities).toHaveBeenCalledWith(false);
  });

  it('two concurrent solution-table loads share one fetch, whatever the id order', async () => {
    const reply = deferred<SolutionComponent[]>();
    client.getSolutionComponents.mockReturnValueOnce(reply.promise);
    const a = loadSolutionEntityNames(['s1', 's2']);
    const b = loadSolutionEntityNames(['s2', 's1']);
    reply.resolve([component('account'), component('contact'), component('account')]);
    expect(await a).toEqual(['account', 'contact']);
    expect(await b).toEqual(['account', 'contact']);
    expect(client.getSolutionComponents).toHaveBeenCalledTimes(1);
  });

  it('concurrent loads of different tables are not merged', async () => {
    client.getEntityAttributes.mockResolvedValue([]);
    await Promise.all([loadEntityAttributes('account'), loadEntityAttributes('contact')]);
    expect(client.getEntityAttributes).toHaveBeenCalledTimes(2);
  });

  it('the same table in two orgs is fetched once per org', async () => {
    client.getEntityAttributes.mockResolvedValue([]);
    const a = loadEntityAttributes('account');
    org.current = `${org.current}-other`;
    const b = loadEntityAttributes('account');
    await Promise.all([a, b]);
    expect(client.getEntityAttributes).toHaveBeenCalledTimes(2);
  });

  it('a later load after the fetch settles is a cache hit', async () => {
    client.getAllEntities.mockResolvedValue([entity('account')]);
    await loadAllEntities();
    await loadAllEntities();
    expect(client.getAllEntities).toHaveBeenCalledTimes(1);
  });
});

describe('loaders: failures are not cached', () => {
  it('a failed attribute load rejects every waiting caller, then the next load retries', async () => {
    const reply = deferred<AttributeMetadata[]>();
    client.getEntityAttributes.mockReturnValueOnce(reply.promise);
    const a = loadEntityAttributes('account');
    const b = loadEntityAttributes('account');
    reply.reject(new Error('throttled'));
    await expect(a).rejects.toThrow('throttled');
    await expect(b).rejects.toThrow('throttled');
    client.getEntityAttributes.mockResolvedValueOnce([attribute('name')]);
    expect(await loadEntityAttributes('account')).toEqual([attribute('name')]);
    expect(client.getEntityAttributes).toHaveBeenCalledTimes(2);
  });

  it('a failed entity-list load is retried on the next call', async () => {
    client.getAllEntities.mockRejectedValueOnce(new Error('offline'));
    await expect(loadAllEntities()).rejects.toThrow('offline');
    client.getAllEntities.mockResolvedValueOnce([entity('account')]);
    expect(await loadAllEntities()).toHaveLength(1);
    expect(client.getAllEntities).toHaveBeenCalledTimes(2);
  });

  it('a failed solution-table load is retried on the next call', async () => {
    client.getSolutionComponents.mockRejectedValueOnce(new Error('offline'));
    await expect(loadSolutionEntityNames(['s1'])).rejects.toThrow('offline');
    client.getSolutionComponents.mockResolvedValueOnce([component('account')]);
    expect(await loadSolutionEntityNames(['s1'])).toEqual(['account']);
    expect(client.getSolutionComponents).toHaveBeenCalledTimes(2);
  });
});

describe('an expired entry being refetched (#334)', () => {
  beforeEach(() => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date('2026-01-01T09:00:00Z'));
  });

  it('callers wait for the refetch instead of getting the expired data', async () => {
    client.getEntityAttributes.mockResolvedValueOnce([attribute('old')]);
    await loadEntityAttributes('account');
    vi.advanceTimersByTime(HOUR + 1);
    const reply = deferred<AttributeMetadata[]>();
    client.getEntityAttributes.mockReturnValueOnce(reply.promise);
    const first = loadEntityAttributes('account');
    // Mid-refetch: the expired data isn't served as fresh.
    expect(metadataCache.getEntityAttributes('account_false')).toBeUndefined();
    const second = loadEntityAttributes('account');
    reply.resolve([attribute('new')]);
    expect(await first).toEqual([attribute('new')]);
    expect(await second).toEqual([attribute('new')]);
    expect(client.getEntityAttributes).toHaveBeenCalledTimes(2);
  });

  it('a failed attribute refetch leaves the entry expired, so the next load retries', async () => {
    client.getEntityAttributes.mockResolvedValueOnce([attribute('old')]);
    await loadEntityAttributes('account');
    vi.advanceTimersByTime(HOUR + 1);
    client.getEntityAttributes.mockRejectedValueOnce(new Error('throttled'));
    await expect(loadEntityAttributes('account')).rejects.toThrow('throttled');
    expect(metadataCache.getEntityAttributes('account_false')).toBeUndefined();
    client.getEntityAttributes.mockResolvedValueOnce([attribute('new')]);
    expect(await loadEntityAttributes('account')).toEqual([attribute('new')]);
    expect(client.getEntityAttributes).toHaveBeenCalledTimes(3);
  });

  it('a failed entity-list refetch leaves the list expired, so the next load retries', async () => {
    client.getAllEntities.mockResolvedValueOnce([entity('old')]);
    await loadAllEntities();
    vi.advanceTimersByTime(HOUR + 1);
    client.getAllEntities.mockRejectedValueOnce(new Error('offline'));
    await expect(loadAllEntities()).rejects.toThrow('offline');
    expect(metadataCache.getAllEntities()).toBeUndefined();
    client.getAllEntities.mockResolvedValueOnce([entity('new')]);
    expect(await loadAllEntities()).toEqual([entity('new')]);
    expect(client.getAllEntities).toHaveBeenCalledTimes(3);
  });

  it("a fresh entry is served at once while another table's refetch is in flight", async () => {
    client.getEntityAttributes.mockResolvedValueOnce([attribute('contactid')]);
    await loadEntityAttributes('contact');
    vi.advanceTimersByTime(HOUR - 1);
    client.getEntityAttributes.mockResolvedValueOnce([attribute('accountid')]);
    await loadEntityAttributes('account');
    vi.advanceTimersByTime(2);
    // contact has expired, account hasn't.
    const reply = deferred<AttributeMetadata[]>();
    client.getEntityAttributes.mockReturnValueOnce(reply.promise);
    const contact = loadEntityAttributes('contact');
    expect(await loadEntityAttributes('account')).toEqual([attribute('accountid')]);
    reply.resolve([attribute('contactid2')]);
    expect(await contact).toEqual([attribute('contactid2')]);
    expect(client.getEntityAttributes).toHaveBeenCalledTimes(3);
  });
});
