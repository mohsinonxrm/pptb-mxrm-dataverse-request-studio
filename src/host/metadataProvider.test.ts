// The metadata provider's table builds (#235): one build per table at a time,
// the live registry behind peekTable, refresh / invalidate (#53), failures not
// cached, alternate keys (#64) and one registry per org (#52).
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type {
  AttributeMetadata,
  EntityKeyMetadata,
  EntityMetadata,
  RelationshipMetadata,
} from './pptbClient';

const org = vi.hoisted(() => ({ current: 'org-a' }));
vi.mock('./activeOrg', () => ({
  activeOrg: () => org.current,
  onActiveOrgChange: () => () => {},
}));

const client = vi.hoisted(() => ({
  getAllEntities: vi.fn(),
  getEntityMetadata: vi.fn(),
  getEntityAttributes: vi.fn(),
  getEntityRelationships: vi.fn(),
  getEntityKeys: vi.fn(),
  getAttributeTypes: vi.fn(),
  getAttributesByTypeCast: vi.fn(),
}));
vi.mock('./pptbClient', async (importOriginal) => ({
  ...(await importOriginal<typeof import('./pptbClient')>()),
  ...client,
}));

import { metadata, subscribeMetadataReset } from './metadataProvider';
import { loadEntityAttributes } from './dataverseMetadata';
import { __subscribeLiveTables } from '../mock/metadata';

type Rels = {
  oneToMany: RelationshipMetadata[];
  manyToOne: RelationshipMetadata[];
  manyToMany: RelationshipMetadata[];
};

const entity = (LogicalName: string, label?: string): EntityMetadata => ({
  LogicalName,
  SchemaName: LogicalName,
  DisplayName: label ? { UserLocalizedLabel: { Label: label } } : undefined,
  EntitySetName: `${LogicalName}s`,
  PrimaryIdAttribute: `${LogicalName}id`,
  PrimaryNameAttribute: 'name',
  MetadataId: `${LogicalName}-meta`,
  ObjectTypeCode: 1,
});

const attribute = (LogicalName: string, AttributeType: string): AttributeMetadata => ({
  LogicalName,
  SchemaName: LogicalName,
  AttributeType,
  MetadataId: LogicalName,
});

const relationship = (r: Partial<RelationshipMetadata> & { SchemaName: string }) =>
  ({
    RelationshipType: 'OneToManyRelationship',
    MetadataId: r.SchemaName,
    ...r,
  }) as RelationshipMetadata;

const key = (
  LogicalName: string,
  KeyAttributes: string[],
  EntityKeyIndexStatus: EntityKeyMetadata['EntityKeyIndexStatus'] = 'Active',
): EntityKeyMetadata => ({
  LogicalName,
  SchemaName: LogicalName,
  KeyAttributes,
  EntityKeyIndexStatus,
  MetadataId: LogicalName,
});

const NO_RELS: Rels = { oneToMany: [], manyToOne: [], manyToMany: [] };

/** Answer every table-build call for `logical` with a minimal table. */
function serveTable(
  logical: string,
  extra: { attrs?: AttributeMetadata[]; rels?: Rels; keys?: EntityKeyMetadata[] } = {},
) {
  client.getEntityMetadata.mockImplementation(async (l: string) => entity(l, `${l} label`));
  client.getEntityAttributes.mockImplementation(async (l: string) =>
    l === logical
      ? (extra.attrs ?? [attribute(`${l}id`, 'Uniqueidentifier'), attribute('name', 'String')])
      : [attribute(`${l}id`, 'Uniqueidentifier')],
  );
  client.getEntityRelationships.mockImplementation(async (l: string) =>
    l === logical ? (extra.rels ?? NO_RELS) : NO_RELS,
  );
  client.getEntityKeys.mockImplementation(async (l: string) =>
    l === logical ? (extra.keys ?? []) : [],
  );
}

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
  // A fresh org per test: the cache, registry and in-flight builds are all
  // per org, so nothing leaks between tests.
  org.current = `org-${++orgSeq}`;
  vi.resetAllMocks();
  vi.spyOn(console, 'error').mockImplementation(() => {});
  vi.spyOn(console, 'warn').mockImplementation(() => {});
});

afterEach(() => {
  vi.restoreAllMocks();
});

describe('getTable', () => {
  it('builds a table from the entity, attributes, relationships and keys', async () => {
    serveTable('account', {
      attrs: [
        attribute('accountid', 'Uniqueidentifier'),
        attribute('name', 'String'),
        attribute('primarycontactid', 'Lookup'),
        attribute('namevirtual', 'Virtual'),
      ],
    });
    const t = await metadata.getTable('account');
    expect(t).toMatchObject({
      logicalName: 'account',
      entitySetName: 'accounts',
      displayName: 'account label',
      primaryKey: 'accountid',
      primaryName: 'name',
      supportsMerge: true,
      alternateKeys: [],
    });
    expect(t?.columns.map((c) => c.logicalName)).toEqual(['accountid', 'name', 'primarycontactid']);
    expect(t?.columns.find((c) => c.logicalName === 'accountid')).toMatchObject({
      attributeType: 'Uniqueidentifier',
      isPrimaryKey: true,
    });
    expect(t?.columns.find((c) => c.logicalName === 'primarycontactid')?.oDataName).toBe(
      '_primarycontactid_value',
    );
    expect(client.getEntityAttributes).toHaveBeenCalledWith('account', false);
    expect(client.getEntityRelationships).toHaveBeenCalledWith('account', false);
  });

  it('marks merge support only for account, contact, lead and incident', async () => {
    serveTable('task');
    expect((await metadata.getTable('task'))?.supportsMerge).toBe(false);
  });

  it('names navigation properties from the side of the relationship this table is on', async () => {
    serveTable('account', {
      rels: {
        manyToOne: [
          relationship({
            SchemaName: 'account_primary_contact',
            ReferencedEntity: 'contact',
            ReferencingAttribute: 'primarycontactid',
            ReferencingEntityNavigationPropertyName: 'primarycontactid',
          }),
          relationship({
            SchemaName: 'account_customer',
            ReferencedEntity: 'contact',
            ReferencingAttribute: 'customerid',
            ReferencingEntityNavigationPropertyName: 'customerid_contact',
          }),
        ],
        oneToMany: [
          relationship({
            SchemaName: 'contact_customer_accounts',
            ReferencingEntity: 'contact',
            ReferencedEntityNavigationPropertyName: 'contact_customer_accounts',
          }),
          relationship({ SchemaName: 'account_tasks', ReferencingEntity: 'task' }),
        ],
        manyToMany: [
          relationship({
            SchemaName: 'accountleads_association',
            Entity1LogicalName: 'account',
            Entity2LogicalName: 'lead',
            Entity1NavigationPropertyName: 'accountleads_association',
            Entity2NavigationPropertyName: 'leadaccounts',
          }),
          relationship({
            SchemaName: 'competitoraccount',
            Entity1LogicalName: 'competitor',
            Entity2LogicalName: 'account',
            Entity1NavigationPropertyName: 'competitoraccounts',
            Entity2NavigationPropertyName: 'accountcompetitors',
            IsValidForAdvancedFind: false,
          }),
        ],
      },
    });
    const navs = (await metadata.getTable('account'))?.navigationProperties ?? [];
    expect(navs).toEqual([
      {
        name: 'primarycontactid',
        relationshipName: 'account_primary_contact',
        cardinality: 'ManyToOne',
        targetEntity: 'contact',
        referencingAttribute: 'primarycontactid',
        isValidForAdvancedFind: true,
      },
      {
        name: 'customerid_contact',
        relationshipName: 'account_customer',
        cardinality: 'ManyToOne',
        targetEntity: 'contact',
        referencingAttribute: 'customerid',
        isValidForAdvancedFind: true,
      },
      {
        name: 'contact_customer_accounts',
        relationshipName: 'contact_customer_accounts',
        cardinality: 'OneToMany',
        targetEntity: 'contact',
        isValidForAdvancedFind: true,
      },
      {
        // No nav property name in the metadata: fall back to the schema name.
        name: 'account_tasks',
        relationshipName: 'account_tasks',
        cardinality: 'OneToMany',
        targetEntity: 'task',
        isValidForAdvancedFind: true,
      },
      {
        name: 'accountleads_association',
        relationshipName: 'accountleads_association',
        cardinality: 'ManyToMany',
        targetEntity: 'lead',
        isValidForAdvancedFind: true,
      },
      {
        name: 'accountcompetitors',
        relationshipName: 'competitoraccount',
        cardinality: 'ManyToMany',
        targetEntity: 'competitor',
        isValidForAdvancedFind: false,
      },
    ]);
  });

  it('keeps only active alternate keys that have columns', async () => {
    serveTable('account', {
      keys: [
        {
          ...key('akey_number', ['name']),
          DisplayName: { UserLocalizedLabel: { Label: 'Number' } },
        },
        key('akey_pending', ['name'], 'Pending'),
        key('akey_failed', ['name'], 'Failed'),
        key('akey_empty', []),
      ],
    });
    expect((await metadata.getTable('account'))?.alternateKeys).toEqual([
      { name: 'akey_number', displayName: 'Number', columns: ['name'] },
    ]);
    expect(client.getAttributeTypes).not.toHaveBeenCalled();
  });

  it('types alternate-key columns the attribute load left out (#64)', async () => {
    serveTable('canvasapp', {
      keys: [key('akey_app', ['name', 'componentstate', 'overwritetime'])],
    });
    client.getAttributeTypes.mockResolvedValue([
      { LogicalName: 'componentstate', AttributeType: 'Picklist' },
      { LogicalName: 'overwritetime', AttributeType: 'DateTime' },
    ]);
    const t = await metadata.getTable('canvasapp');
    expect(client.getAttributeTypes).toHaveBeenCalledWith('canvasapp', [
      'componentstate',
      'overwritetime',
    ]);
    expect(t?.alternateKeys?.[0].columnTypes).toEqual({
      componentstate: 'Picklist',
      overwritetime: 'DateTime',
    });
  });

  it('types those key columns through attributeTypeOf, like every other column (#318)', async () => {
    serveTable('new_doc', { keys: [key('akey_doc', ['new_ref', 'new_file', 'new_virtual'])] });
    client.getAttributeTypes.mockResolvedValue([
      { LogicalName: 'new_ref', AttributeType: 'String' },
      {
        LogicalName: 'new_file',
        AttributeType: 'Virtual',
        AttributeTypeName: { Value: 'FileType' },
      },
      {
        LogicalName: 'new_virtual',
        AttributeType: 'Virtual',
        AttributeTypeName: { Value: 'MultiSelectPicklistType' },
      },
    ]);
    const t = await metadata.getTable('new_doc');
    expect(t?.alternateKeys?.[0].columnTypes).toEqual({ new_ref: 'String', new_file: 'File' });
  });

  it('still builds the table when typing the missing key columns fails', async () => {
    serveTable('canvasapp', { keys: [key('akey_app', ['componentstate'])] });
    client.getAttributeTypes.mockRejectedValue(new Error('throttled'));
    const t = await metadata.getTable('canvasapp');
    expect(t?.alternateKeys).toEqual([
      { name: 'akey_app', displayName: 'akey_app', columns: ['componentstate'] },
    ]);
  });

  it('two concurrent reads of the same table share one build', async () => {
    serveTable('account');
    const entityReply = deferred<EntityMetadata>();
    client.getEntityMetadata.mockReturnValueOnce(entityReply.promise);
    const a = metadata.getTable('account');
    const b = metadata.getTable('account');
    entityReply.resolve(entity('account'));
    const [ta, tb] = await Promise.all([a, b]);
    expect(ta).toBe(tb);
    expect(client.getEntityMetadata).toHaveBeenCalledTimes(1);
    expect(client.getEntityAttributes).toHaveBeenCalledTimes(1);
    expect(client.getEntityRelationships).toHaveBeenCalledTimes(1);
    expect(client.getEntityKeys).toHaveBeenCalledTimes(1);
  });

  it('a read after the build lands is served from the registry without fetching', async () => {
    serveTable('account');
    const first = await metadata.getTable('account');
    vi.clearAllMocks();
    expect(await metadata.getTable('account')).toBe(first);
    expect(client.getEntityMetadata).not.toHaveBeenCalled();
  });

  it('a failed build resolves to undefined and registers nothing', async () => {
    serveTable('account');
    client.getEntityRelationships.mockRejectedValueOnce(new Error('throttled'));
    expect(await metadata.getTable('account')).toBeUndefined();
    expect(metadata.peekTable('account')).toBeUndefined();
  });

  it('a failed build is not cached: the next read builds again', async () => {
    serveTable('account');
    client.getEntityMetadata.mockRejectedValueOnce(new Error('offline'));
    expect(await metadata.getTable('account')).toBeUndefined();
    const t = await metadata.getTable('account');
    expect(t?.logicalName).toBe('account');
    expect(client.getEntityMetadata).toHaveBeenCalledTimes(2);
  });
});

describe('peekTable', () => {
  it('returns nothing until the table is built, and never fetches', async () => {
    serveTable('account');
    expect(metadata.peekTable('account')).toBeUndefined();
    expect(client.getEntityMetadata).not.toHaveBeenCalled();
    const pending = metadata.getTable('account');
    expect(metadata.peekTable('account')).toBeUndefined();
    const built = await pending;
    expect(metadata.peekTable('account')).toBe(built);
  });
});

describe('listEntities', () => {
  it('maps the entity list to picker items', async () => {
    client.getAllEntities.mockResolvedValue([
      entity('account', 'Account'),
      { ...entity('activitypointer'), IsValidForAdvancedFind: false },
    ]);
    expect(await metadata.listEntities()).toEqual([
      {
        logicalName: 'account',
        displayName: 'Account',
        entitySetName: 'accounts',
        metadataId: 'account-meta',
        isValidForAdvancedFind: true,
      },
      {
        logicalName: 'activitypointer',
        displayName: 'activitypointer',
        entitySetName: 'activitypointers',
        metadataId: 'activitypointer-meta',
        isValidForAdvancedFind: false,
      },
    ]);
  });

  it('fetches the entity list once for repeat calls', async () => {
    client.getAllEntities.mockResolvedValue([entity('account')]);
    await metadata.listEntities();
    await metadata.listEntities();
    expect(client.getAllEntities).toHaveBeenCalledTimes(1);
  });
});

describe('invalidateAll (#53)', () => {
  it('empties the registry at once, so the next read refetches', async () => {
    serveTable('account');
    await metadata.getTable('account');
    metadata.invalidateAll();
    expect(metadata.peekTable('account')).toBeUndefined();
    await metadata.getTable('account');
    expect(client.getEntityMetadata).toHaveBeenCalledTimes(2);
    expect(client.getEntityAttributes).toHaveBeenCalledTimes(2);
  });

  it('drops the cached entity list too', async () => {
    client.getAllEntities.mockResolvedValue([entity('account')]);
    await metadata.listEntities();
    metadata.invalidateAll();
    await metadata.listEntities();
    expect(client.getAllEntities).toHaveBeenCalledTimes(2);
  });

  it('tells reset subscribers, until they unsubscribe', () => {
    const listener = vi.fn();
    const off = subscribeMetadataReset(listener);
    metadata.invalidateAll();
    expect(listener).toHaveBeenCalledTimes(1);
    off();
    metadata.invalidateAll();
    expect(listener).toHaveBeenCalledTimes(1);
  });
});

describe('refreshAll', () => {
  it('rebuilds every registered table, keeping the old data visible until the new lands', async () => {
    serveTable('account');
    const old = await metadata.getTable('account');
    const entityReply = deferred<EntityMetadata>();
    client.getEntityMetadata.mockReturnValueOnce(entityReply.promise);

    const refresh = metadata.refreshAll();
    await vi.waitFor(() => expect(client.getEntityMetadata).toHaveBeenCalledTimes(2));
    expect(metadata.peekTable('account')).toBe(old);

    entityReply.resolve(entity('account', 'Account (renamed)'));
    await refresh;
    const fresh = metadata.peekTable('account');
    expect(fresh).not.toBe(old);
    expect(fresh?.displayName).toBe('Account (renamed)');
    expect(client.getEntityAttributes).toHaveBeenCalledTimes(2);
  });

  it('waits for a build already in flight, then refetches it', async () => {
    serveTable('account');
    const firstReply = deferred<EntityMetadata>();
    client.getEntityMetadata.mockReturnValueOnce(firstReply.promise);
    const build = metadata.getTable('account');

    let refreshed = false;
    const refresh = metadata.refreshAll().then(() => (refreshed = true));
    // A full macrotask: a refresh that didn't wait would have settled by now.
    await new Promise((r) => setTimeout(r, 0));
    expect(refreshed).toBe(false);

    firstReply.resolve(entity('account', 'before refresh'));
    await build;
    await refresh;
    expect(metadata.peekTable('account')?.displayName).toBe('account label');
    expect(client.getEntityMetadata).toHaveBeenCalledTimes(2);
  });

  it('tells reset subscribers', async () => {
    const listener = vi.fn();
    const off = subscribeMetadataReset(listener);
    await metadata.refreshAll();
    expect(listener).toHaveBeenCalledTimes(1);
    off();
  });

  it('keeps going when one table fails to rebuild', async () => {
    serveTable('account');
    await metadata.getTable('account');
    await metadata.getTable('contact');
    client.getEntityMetadata.mockImplementation(async (l: string) => {
      if (l === 'contact') throw new Error('throttled');
      return entity(l, 'Account (fresh)');
    });
    await expect(metadata.refreshAll()).resolves.toBeUndefined();
    expect(metadata.peekTable('account')?.displayName).toBe('Account (fresh)');
    // The failed table keeps its last good build.
    expect(metadata.peekTable('contact')?.logicalName).toBe('contact');
  });
});

describe('invalidateTable', () => {
  it("drops the table's cached metadata, leaving other tables alone", async () => {
    serveTable('account');
    await metadata.getTable('account');
    await metadata.getTable('contact');
    metadata.invalidateTable('account');
    vi.clearAllMocks();
    await loadEntityAttributes('account');
    await loadEntityAttributes('contact');
    expect(client.getEntityAttributes).toHaveBeenCalledTimes(1);
    expect(client.getEntityAttributes).toHaveBeenCalledWith('account', false);
  });

  it('the next read rebuilds the table, and other tables stay registered (#335)', async () => {
    serveTable('account');
    const old = await metadata.getTable('account');
    const contact = await metadata.getTable('contact');
    metadata.invalidateTable('account');
    expect(metadata.peekTable('account')).toBeUndefined();
    expect(metadata.peekTable('contact')).toBe(contact);
    vi.clearAllMocks();
    serveTable('account');
    const fresh = await metadata.getTable('account');
    expect(fresh).toBeDefined();
    expect(fresh).not.toBe(old);
    expect(client.getEntityMetadata).toHaveBeenCalledWith('account');
  });

  it("a build still running when the table is invalidated doesn't put it back (#335)", async () => {
    serveTable('account');
    const reply = deferred<EntityMetadata>();
    client.getEntityMetadata.mockReturnValueOnce(reply.promise);
    const stale = metadata.getTable('account');
    metadata.invalidateTable('account');
    reply.resolve(entity('account', 'stale'));
    await stale;
    expect(metadata.peekTable('account')).toBeUndefined();
    const fresh = await metadata.getTable('account');
    expect(fresh?.displayName).toBe('account label');
  });

  it('tells the registry listeners, so open editors refetch (#335)', async () => {
    serveTable('account');
    await metadata.getTable('account');
    const seen = vi.fn();
    const off = __subscribeLiveTables(seen);
    metadata.invalidateTable('account');
    off();
    expect(seen).toHaveBeenCalled();
  });
});

describe('one registry per org (#52)', () => {
  it("never shows another org's table, and keeps it for switching back", async () => {
    serveTable('account');
    const first = org.current;
    const built = await metadata.getTable('account');
    org.current = `${first}-other`;
    expect(metadata.peekTable('account')).toBeUndefined();
    await metadata.getTable('account');
    expect(client.getEntityMetadata).toHaveBeenCalledTimes(2);
    org.current = first;
    expect(metadata.peekTable('account')).toBe(built);
  });

  it('the same table in two orgs is built separately, even at the same time', async () => {
    serveTable('account');
    const a = metadata.getTable('account');
    org.current = `${org.current}-other`;
    const b = metadata.getTable('account');
    const [ta, tb] = await Promise.all([a, b]);
    expect(ta).not.toBe(tb);
    expect(client.getEntityMetadata).toHaveBeenCalledTimes(2);
  });

  it('a build that lands after a switch is filed under the org it started in', async () => {
    serveTable('account');
    const entityReply = deferred<EntityMetadata>();
    client.getEntityMetadata.mockReturnValueOnce(entityReply.promise);
    const first = org.current;
    const build = metadata.getTable('account');
    org.current = `${first}-other`;
    entityReply.resolve(entity('account'));
    await build;
    expect(metadata.peekTable('account')).toBeUndefined();
    org.current = first;
    expect(metadata.peekTable('account')?.logicalName).toBe('account');
  });

  it("invalidateAll empties only the active org's registry", async () => {
    serveTable('account');
    const first = org.current;
    await metadata.getTable('account');
    org.current = `${first}-other`;
    metadata.invalidateAll();
    org.current = first;
    expect(metadata.peekTable('account')).toBeDefined();
  });
});
