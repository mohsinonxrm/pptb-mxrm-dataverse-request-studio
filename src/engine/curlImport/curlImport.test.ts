// cURL import (#271): which mode a command opens, what it carries over, and
// what it refuses.
import { beforeEach, describe, expect, it, vi } from 'vitest';
import { importCurl, parsePrefer, type CurlImportContext } from './curlImport';
import { account, contact, registerFixtureTables } from '../../test/fixtures';
import type { CsdlAction } from '../../mock/actionsCsdl';
import { encodeParamForBody } from '../executeBuilders';
import { __clearLiveTables, __registerLiveTables } from '../../mock/metadata';

vi.mock('../../mock/environment', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../../mock/environment')>()),
  ENV: {
    name: 'Test',
    host: 'contoso.crm.dynamics.com',
    apiBase: '/api/data/v9.2',
    user: '',
  },
}));

const ORG = 'https://contoso.crm.dynamics.com/api/data/v9.2';
const ID = '00000000-0000-0000-0000-0000000000a1';

const action = (name: string, extra: Partial<CsdlAction>): CsdlAction =>
  ({
    name,
    kind: 'Action',
    namespace: 'Microsoft.Dynamics.CRM',
    binding: { kind: 'unbound' },
    parameters: [],
    returnType: { kind: 'none' },
    source: 'oob',
    ...extra,
  }) as CsdlAction;

const ACTIONS: Record<string, CsdlAction> = {
  // Shaped like the real one; the fixture's account stands in for opportunityclose.
  WinOpportunity: action('WinOpportunity', {
    parameters: [
      { name: 'Status', type: 'OptionSetValue', required: true },
      { name: 'OpportunityClose', type: 'EntitySpecific', entityType: 'account', required: true },
    ],
  }),
  new_Route: action('new_Route', {
    parameters: [
      { name: 'Target', type: 'EntityReference', required: true },
      { name: 'Members', type: 'EntityCollection', entityType: 'contact', required: false },
      { name: 'Note', type: 'Edm.String', required: false },
    ],
  }),
  new_Ping: action('new_Ping', { source: 'custom-api' }),
  new_Flag: action('new_Flag', {
    source: 'custom-action',
    binding: { kind: 'entity', entityType: 'account' },
  }),
  WhoAmI: action('WhoAmI', { kind: 'Function' }),
};

const ctx: CurlImportContext = {
  host: 'contoso.crm.dynamics.com',
  resolveEntitySet: (set) => ({ accounts: 'account', contacts: 'contact' })[set],
  loadTable: async (logical) => ({ account, contact })[logical as 'account'],
  findAction: async (name) => ACTIONS[name],
};

const run = (cmd: string) => importCurl(cmd, ctx);

/** A bash `$'…'` word for `s`: backslash first, then quotes and line breaks. */
const ansiC = (s: string) =>
  `$'${s.replace(/\\/g, '\\\\').replace(/'/g, "\\'").replace(/\r/g, '\\r').replace(/\n/g, '\\n')}'`;

beforeEach(() => registerFixtureTables());

describe('reads', () => {
  it('a collection GET opens Retrieve Multiple with the query parsed', async () => {
    const r = await run(
      `curl '${ORG}/accounts?$select=name&$filter=numberofemployees gt 10&$top=5' -H 'Prefer: odata.include-annotations="OData.Community.Display.V1.FormattedValue"'`,
    );
    expect(r.errors).toEqual([]);
    expect(r.request).toMatchObject({
      modeId: 'retrieve-multiple',
      parsed: { table: 'account', select: ['name'], top: 5 },
      prefer: { formattedValues: true, allAnnotations: false },
    });
    if (r.request?.modeId !== 'retrieve-multiple') throw new Error();
    expect(r.request.parsed.filter.rules).toHaveLength(1);
  });

  it('a /$count GET opens Count only', async () => {
    const r = await run(`curl ${ORG}/accounts/\\$count`);
    expect(r.request).toMatchObject({ modeId: 'retrieve-multiple', parsed: { countOnly: true } });
  });

  it("passes on the parser's errors", async () => {
    const r = await run(`curl '${ORG}/accounts?$select=nosuchcolumn'`);
    expect(r.request).toBeUndefined();
    expect(r.errors.join(' ')).toMatch(/nosuchcolumn/);
  });

  it('a row GET opens Retrieve Single and says what it leaves out', async () => {
    const r = await run(`curl '${ORG}/accounts({${ID.toUpperCase()}})?$select=name&$top=1'`);
    expect(r.request).toMatchObject({
      modeId: 'retrieve-single',
      table: 'account',
      recordId: ID,
      select: ['name'],
    });
    expect(r.warnings.join(' ')).toMatch(/Left out \$top: Retrieve Single sends only/);
  });
});

describe('writes', () => {
  it('POST to a set opens Create with the JSON body, Prefer and $select', async () => {
    const r = await run(
      `curl '${ORG}/accounts?$select=name' -H 'Content-Type: application/json' -H 'Prefer: return=representation' -H 'MSCRM.SuppressDuplicateDetection: false' --data-raw '{"name":"A","primarycontactid@odata.bind":"/contacts(${ID})"}'`,
    );
    expect(r.errors).toEqual([]);
    expect(r.request).toMatchObject({
      modeId: 'create',
      table: 'account',
      rawBody: `{\n  "name": "A",\n  "primarycontactid@odata.bind": "/contacts(${ID})"\n}`,
      prefer: { returnRepresentation: true, formattedValues: false },
      returnSelect: ['name'],
      headers: [
        { name: 'Content-Type', value: 'application/json' },
        { name: 'MSCRM.SuppressDuplicateDetection', value: 'false' },
      ],
    });
  });

  it('PATCH with If-Match opens Update; without, Upsert', async () => {
    const body = `-d '{"name":"B"}'`;
    expect(
      (await run(`curl -X PATCH '${ORG}/accounts(${ID})' -H 'If-Match: *' ${body}`)).request,
    ).toMatchObject({
      modeId: 'update',
      recordId: ID,
      concurrency: { kind: 'update-only' },
    });
    expect(
      (await run(`curl -X PATCH '${ORG}/accounts(${ID})' -H 'If-Match: W/"123"' ${body}`)).request,
    ).toMatchObject({ modeId: 'update', concurrency: { kind: 'etag', etag: 'W/"123"' } });
    expect((await run(`curl -X PATCH '${ORG}/accounts(${ID})' ${body}`)).request).toMatchObject({
      modeId: 'upsert',
      concurrency: { kind: 'none' },
    });
    expect(
      (await run(`curl -X PATCH '${ORG}/accounts(${ID})' -H 'If-None-Match: *' ${body}`)).request,
    ).toMatchObject({ modeId: 'upsert', concurrency: { kind: 'create-only' } });
  });

  it('reads the bypass headers into the Bypass pane', async () => {
    const r = await run(
      `curl -X PATCH '${ORG}/accounts(${ID})' -H 'If-Match: *' -H 'MSCRM.BypassBusinessLogicExecution: CustomSync,CustomAsync' -H 'MSCRM.SuppressCallbackRegistrationExpanderJob: true' -d '{}'`,
    );
    expect(r.request).toMatchObject({
      bypass: { businessLogic: 'both', suppressFlows: true },
      headers: [],
    });
  });

  it('DELETE opens Delete, for a row or one column', async () => {
    const star = await run(`curl -X DELETE '${ORG}/accounts(${ID})' -H 'If-Match: *'`);
    expect(star.request).toMatchObject({
      modeId: 'delete',
      recordId: ID,
      scope: { kind: 'whole-row' },
      // Delete's precondition pane offers None or an ETag only.
      concurrency: { kind: 'none' },
    });
    expect(star.notes.join(' ')).toMatch(/Left out If-Match: \*/);
    expect(
      (await run(`curl -X DELETE '${ORG}/accounts(${ID})' -H 'If-Match: W/"9"'`)).request,
    ).toMatchObject({ concurrency: { kind: 'etag', etag: 'W/"9"' } });
    expect((await run(`curl -X DELETE '${ORG}/accounts(${ID})/name'`)).request).toMatchObject({
      modeId: 'delete',
      scope: { kind: 'single-property', column: 'name' },
      concurrency: { kind: 'none' },
    });
    expect((await run(`curl -X DELETE '${ORG}/accounts(${ID})/nope'`)).errors.join()).toMatch(
      /no column named "nope"/,
    );
  });

  it('refuses a write without a usable body or row', async () => {
    expect((await run(`curl -X PATCH '${ORG}/accounts(${ID})'`)).errors.join()).toMatch(
      /needs a JSON body/,
    );
    expect((await run(`curl '${ORG}/accounts' -d '[1]'`)).errors.join()).toMatch(/JSON object/);
    expect(
      (
        await run(
          `curl '${ORG}/accounts' -d 'a=1' -H 'Content-Type: application/x-www-form-urlencoded'`,
        )
      ).errors.join(),
    ).toMatch(/not JSON/);
    expect((await run(`curl -X DELETE '${ORG}/accounts'`)).errors.join()).toMatch(/needs a row/);
  });
});

describe('actions', () => {
  it('an unbound POST opens the mode for its category, with the body as parameters', async () => {
    expect((await run(`curl '${ORG}/WinOpportunity' -d '{"Status":3}'`)).request).toEqual({
      modeId: 'exec-action',
      actionName: 'WinOpportunity',
      boundRecordId: null,
      paramValues: { Status: 3 },
      headers: [],
    });
    expect((await run(`curl -X POST '${ORG}/new_Ping'`)).request).toMatchObject({
      modeId: 'exec-customapi',
      paramValues: {},
    });
  });

  it('a bound POST checks the binding', async () => {
    expect(
      (await run(`curl -X POST '${ORG}/accounts(${ID})/Microsoft.Dynamics.CRM.new_Flag'`)).request,
    ).toMatchObject({ modeId: 'exec-customaction', boundRecordId: ID });
    expect(
      (
        await run(`curl -X POST '${ORG}/contacts(${ID})/Microsoft.Dynamics.CRM.new_Flag'`)
      ).errors.join(),
    ).toMatch(/isn't bound to contact/);
    expect((await run(`curl -X POST '${ORG}/new_Flag'`)).errors.join()).toMatch(
      /is bound to account/,
    );
  });

  it('refuses functions and unknown names', async () => {
    expect((await run(`curl '${ORG}/WhoAmI'`)).errors.join()).toMatch(/doesn't open functions yet/);
    expect((await run(`curl -X POST '${ORG}/WhoAmI'`)).errors.join()).toMatch(
      /No action named WhoAmI/,
    );
    expect((await run(`curl -X POST '${ORG}/nothing'`)).errors.join()).toMatch(
      /No action named nothing/,
    );
  });

  // The form stores row references and new rows in its own shapes; the
  // encoder must turn them back into exactly the JSON the command sent.
  const roundTrip = async (name: string, wire: Record<string, unknown>) => {
    // POSIX single quotes: an apostrophe is written '\''.
    const quoted = `'${JSON.stringify(wire).replace(/'/g, "'\\''")}'`;
    const r = await run(`curl -X POST '${ORG}/${name}' --data-raw ${quoted}`);
    expect(r.errors).toEqual([]);
    if (!r.request || !('paramValues' in r.request)) throw new Error('not an action');
    const sent = Object.fromEntries(
      ACTIONS[name].parameters
        .filter((p) => r.request && 'paramValues' in r.request && p.name in r.request.paramValues)
        .map((p) => [
          p.name,
          encodeParamForBody(
            p,
            (r.request as { paramValues: Record<string, unknown> }).paramValues[p.name],
          ),
        ]),
    );
    expect(sent).toEqual(wire);
    return r.request.paramValues;
  };

  it('reads row references, collections and new rows into the form', async () => {
    expect(
      await roundTrip('new_Route', {
        Target: { '@odata.type': 'Microsoft.Dynamics.CRM.account', accountid: ID },
        Members: [{ '@odata.type': 'Microsoft.Dynamics.CRM.contact', contactid: ID }],
        Note: "it's",
      }),
    ).toEqual({
      Target: { id: ID, entityType: 'account' },
      Members: [{ id: ID, entityType: 'contact' }],
      Note: "it's",
    });
    expect(
      await roundTrip('WinOpportunity', {
        Status: 3,
        OpportunityClose: {
          '@odata.type': 'Microsoft.Dynamics.CRM.account',
          name: 'Won',
          sample_tags: '1,2',
          'primarycontactid@odata.bind': `/contacts(${ID})`,
        },
      }),
    ).toMatchObject({
      OpportunityClose: {
        name: 'Won',
        sample_tags: [1, 2],
        primarycontactid: { id: ID, targetEntity: 'contact' },
      },
    });
    // Just an id: a reference to an existing row.
    expect(
      await roundTrip('WinOpportunity', {
        OpportunityClose: { '@odata.type': 'Microsoft.Dynamics.CRM.account', accountid: ID },
      }),
    ).toEqual({ OpportunityClose: { '@ref': ID } });
  });

  it('reads an @odata.id reference', async () => {
    const r = await run(
      `curl -X POST '${ORG}/new_Route' -d '{"Target":{"@odata.id":"${ORG}/accounts(${ID})"}}'`,
    );
    expect(r.request).toMatchObject({ paramValues: { Target: { id: ID, entityType: 'account' } } });
  });

  it("keeps a new row's lookup when the target table's metadata isn't loaded", async () => {
    // Found live: WinOpportunity's opportunityid@odata.bind was dropped when
    // the opportunity table hadn't been loaded.
    __clearLiveTables();
    __registerLiveTables([account]);
    await roundTrip('WinOpportunity', {
      OpportunityClose: {
        '@odata.type': 'Microsoft.Dynamics.CRM.account',
        'primarycontactid@odata.bind': `/contacts(${ID})`,
      },
    });
  });

  it("refuses parameters it can't read, instead of sending null", async () => {
    const errs = async (body: string) =>
      (await run(`curl -X POST '${ORG}/new_Route' -d '${body}'`)).errors.join(' ');
    expect(await errs('{"Target":{"name":"x"}}')).toMatch(/Target: expected a row reference/);
    expect(await errs('{"Target":"nope"}')).toMatch(/Target: expected a row reference/);
    expect(await errs('{"Members":{}}')).toMatch(/Members: expected an array/);
    expect(await errs('{"Note":5}')).toMatch(/Note: expected a string/);
    expect(await errs('{"Nope":1}')).toMatch(/new_Route has no parameter named "Nope"/);
    expect(
      (
        await run(`curl -X POST '${ORG}/WinOpportunity' -d '{"OpportunityClose":{"nosuch":1}}'`)
      ).errors.join(),
    ).toMatch(/account has no column "nosuch"/);
  });
});

describe('$batch', () => {
  const CT = 'multipart/mixed; boundary=batch_1';
  const body = [
    '--batch_1',
    'Content-Type: multipart/mixed; boundary=changeset_1',
    '',
    '--changeset_1',
    'Content-Type: application/http',
    'Content-Transfer-Encoding: binary',
    'Content-ID: 7',
    '',
    `POST ${ORG}/accounts HTTP/1.1`,
    'Content-Type: application/json',
    'Authorization: Bearer secret',
    '',
    '{"name":"A"}',
    '--changeset_1--',
    '--batch_1',
    'Content-Type: application/http',
    '',
    'GET /api/data/v9.2/accounts?$top=1 HTTP/1.1',
    '',
    '',
    '--batch_1',
    'Content-Type: application/http',
    '',
    'DELETE accounts(1) HTTP/1.1',
    '',
    '',
    '--batch_1--',
    '',
  ].join('\r\n');
  const cmd = (extra = '') =>
    `curl -X POST '${ORG}/$batch' -H 'Content-Type: ${CT}' ${extra} --data-binary ${ansiC(body)}`;

  it('opens $batch with change sets and independent requests grouped as sent', async () => {
    const r = await run(cmd(`-H 'Prefer: odata.continue-on-error'`));
    expect(r.errors).toEqual([]);
    if (r.request?.modeId !== 'batch') throw new Error('not batch');
    expect(r.request.continueOnError).toBe(true);
    expect(
      r.request.groups.map((g) => [g.atomic, g.items.map((i) => `${i.method} ${i.url}`)]),
    ).toEqual([
      [true, [`POST ${ORG}/accounts`]],
      [false, ['GET /api/data/v9.2/accounts?$top=1', 'DELETE accounts(1)']],
    ]);
    const first = r.request.groups[0].items[0];
    expect(first.op).toBe('custom');
    expect(first.body).toBe('{"name":"A"}');
    expect(first.headers).toEqual([
      { name: 'content-type', value: 'application/json', enabled: true },
    ]);
    expect(r.notes.join(' ')).toMatch(/Left out authorization from the batch's requests/);
    expect(r.warnings.join(' ')).toMatch(/numbers Content-IDs 1, 2, 3/);
    expect(JSON.stringify(r)).not.toContain('secret');
  });

  it("refuses a body that isn't multipart", async () => {
    const r = await run(`curl -X POST '${ORG}/$batch' -d 'nope'`);
    expect(r.errors.join()).toMatch(/\$batch body can't be read/);
  });
});

describe('untrusted input', () => {
  it('drops credentials and browser headers, and says so', async () => {
    const r = await run(
      `curl '${ORG}/accounts' -H 'Authorization: Bearer secret' -H 'Cookie: a=b' -H 'sec-ch-ua: x' -H 'User-Agent: y' -H 'MSCRMCallerID: ${ID}'`,
    );
    expect(r.request).toMatchObject({ headers: [{ name: 'MSCRMCallerID', value: ID }] });
    expect(r.notes).toEqual([
      'Left out Authorization: ToolBox signs you in.',
      'Left out Cookie: ToolBox signs you in.',
      'Left out browser headers: sec-ch-ua, User-Agent.',
    ]);
    expect(JSON.stringify(r)).not.toContain('secret');
  });

  it.each([
    [`curl http://contoso.crm.dynamics.com/api/data/v9.2/accounts`, /must use https/],
    [`curl file:///etc/passwd`, /isn't a Dataverse Web API URL/],
    [`curl https://contoso.crm.dynamics.com/accounts`, /isn't under \/api\/data\/v9\.x/],
    [`curl '${ORG}/accounts/../systemusers'`, /"\." or "\.\." segments/],
    [`curl '${ORG}/nosuchset(${ID})'`, /No table with the entity set name/],
    [`curl '${ORG}/accounts(not-a-guid)'`, /isn't a row id/],
    [`curl '${ORG}/accounts(accountnumber=%27A1%27)'`, /alternate key/],
    [
      `curl -X POST '${ORG}/accounts(${ID})/contact_customer_accounts/$ref' -d '{}'`,
      /associations/,
    ],
    [`curl -X OPTIONS '${ORG}/accounts'`, /isn't a Dataverse Web API method/],
    [`curl '${ORG}/accounts' -H 'Bad Name: 1'`, /isn't a header name/],
  ])('%s', async (cmd, re) => {
    const r = await run(cmd);
    expect(r.request).toBeUndefined();
    expect(r.errors.join(' ')).toMatch(re);
  });

  it('warns when the command is for another environment', async () => {
    const r = await run(`curl https://other.crm4.dynamics.com/api/data/v9.2/accounts`);
    expect(r.request?.modeId).toBe('retrieve-multiple');
    expect(r.warnings.join()).toMatch(/for other\.crm4\.dynamics\.com; it opens against .*contoso/);
  });
});

describe('parsePrefer', () => {
  it('reads the switches and lists what it does not know', () => {
    expect(
      parsePrefer(
        'odata.include-annotations="OData.Community.Display.V1.FormattedValue,Microsoft.Dynamics.CRM.lookuplogicalname,x.y", return=representation, odata.maxpagesize=50, respond-async',
      ),
    ).toEqual({
      spec: {
        maxpagesize: 50,
        formattedValues: true,
        lookupLogicalNames: true,
        associatedNavProperty: false,
        totalRecordCount: false,
        allAnnotations: false,
        returnRepresentation: true,
      },
      continueOnError: false,
      unknown: ['odata.include-annotations="x.y"', 'respond-async'],
    });
    expect(parsePrefer('odata.include-annotations="*"').spec.allAnnotations).toBe(true);
  });
});

describe("DRS's own cURL export", () => {
  it('imports back as the request it describes', async () => {
    const { generateCode } = await import('../codeGenerators');
    const curl = generateCode('curl', {
      method: 'PATCH',
      built: {
        relativeUrl: `/api/data/v9.2/accounts(${ID})?$select=name`,
        relativeNoBase: `/accounts(${ID})?$select=name`,
        bytes: 0,
        queryParts: [{ key: '$select', value: 'name' }],
        entitySet: 'accounts',
        entityLogical: 'account',
      },
      headers: {
        'If-Match': '*',
        'Content-Type': 'application/json; charset=utf-8',
        Prefer: 'return=representation',
        Authorization: 'Bearer secret',
      },
      body: { name: 'O\'Brien "q" $HOME', revenue: 1.5 },
    });
    const r = await importCurl(curl, ctx);
    expect(r.errors).toEqual([]);
    expect(r.request).toMatchObject({
      modeId: 'update',
      recordId: ID,
      concurrency: { kind: 'update-only' },
      prefer: { returnRepresentation: true },
      returnSelect: ['name'],
      headers: [{ name: 'Content-Type', value: 'application/json; charset=utf-8' }],
    });
    if (r.request?.modeId !== 'update') throw new Error();
    expect(JSON.parse(r.request.rawBody)).toEqual({ name: 'O\'Brien "q" $HOME', revenue: 1.5 });
  });
});

describe('nothing changes quietly (review of #271)', () => {
  it('keeps bypass headers the Bypass pane cannot show as header rows', async () => {
    const r = await run(
      `curl -X PATCH '${ORG}/accounts(${ID})' -H 'If-Match: *' -H 'MSCRM.BypassBusinessLogicExecution: CustomAsync' -H 'MSCRM.BypassBusinessLogicExecutionStepIds: ${ID}' -d '{}'`,
    );
    expect(r.request).toMatchObject({
      bypass: { businessLogic: 'none' },
      headers: [
        { name: 'MSCRM.BypassBusinessLogicExecution', value: 'CustomAsync' },
        { name: 'MSCRM.BypassBusinessLogicExecutionStepIds', value: ID },
      ],
    });
    expect(r.warnings.join()).toMatch(/Kept the bypass headers as header rows/);
    const steps = await run(
      `curl -X PATCH '${ORG}/accounts(${ID})' -H 'If-Match: *' -H 'MSCRM.BypassBusinessLogicExecutionStepIds: ${ID.toUpperCase()}' -d '{}'`,
    );
    expect(steps.request).toMatchObject({
      bypass: { businessLogic: 'steps', stepIds: [ID] },
      headers: [],
    });
  });

  it('Retrieve Single keeps If-None-Match and says what else it leaves out', async () => {
    const r = await run(
      `curl '${ORG}/accounts(${ID})' -H 'If-None-Match: W/"5"' -H 'Prefer: odata.maxpagesize=5'`,
    );
    expect(r.request).toMatchObject({ headers: [{ name: 'If-None-Match', value: 'W/"5"' }] });
    expect(r.warnings.join()).toMatch(/Left out Prefer: odata.maxpagesize=5: only queries page/);
  });

  it('says when a Create leaves out preconditions or a page size', async () => {
    const r = await run(
      `curl '${ORG}/accounts' -H 'If-Match: *' -H 'Prefer: odata.maxpagesize=5' -d '{}'`,
    );
    expect(r.warnings.join(' ')).toMatch(
      /Left out If-Match: \*: Create doesn't send one.*maxpagesize/,
    );
  });

  it('reads a lookup in $select back to its column', async () => {
    const r = await run(`curl '${ORG}/accounts?$select=name,_primarycontactid_value' -d '{}'`);
    expect(r.request).toMatchObject({ returnSelect: ['name', 'primarycontactid'] });
  });

  it('$batch says what it leaves out and flags other hosts and impersonation', async () => {
    const body = [
      '--b',
      'Content-Type: application/http',
      '',
      `GET https://evil.example.com/api/data/v9.2/accounts HTTP/1.1`,
      `MSCRMCallerID: ${ID}`,
      '',
      '',
      '--b--',
      '',
    ].join('\r\n');
    const r = await run(
      `curl -X POST '${ORG}/$batch' -H 'Content-Type: multipart/mixed; boundary=b' -H 'MSCRM.BypassCustomPluginExecution: true' -H 'Prefer: return=representation' --data-binary ${ansiC(body)}`,
    );
    const w = r.warnings.join(' | ');
    expect(w).toMatch(/batch-level Prefer preferences/);
    expect(w).toMatch(/Left out MSCRM.BypassCustomPluginExecution on the batch itself/);
    expect(w).toMatch(/Request 1 is for evil\.example\.com/);
    expect(w).toMatch(/Request 1 runs as another user \(mscrmcallerid\)/);
  });
});
