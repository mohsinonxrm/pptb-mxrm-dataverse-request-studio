// #248: Custom APIs and Custom Actions carry their solution's publisher.
// A file of its own so the provider's per-org caches start empty.
import { beforeEach, describe, expect, it, vi } from 'vitest';

const getCSDLDocument = vi.hoisted(() => vi.fn<() => Promise<string>>());
const queryData = vi.hoisted(() => vi.fn<(path: string) => Promise<unknown>>());
vi.mock('./pptbBridge', async (importOriginal) => ({
  ...(await importOriginal<typeof import('./pptbBridge')>()),
  dvHost: { metadata: { getCSDLDocument }, queryData },
}));

import { fetchCsdlActions } from './csdlProvider';

const CSDL = `<?xml version="1.0" encoding="utf-8"?>
<edmx:Edmx Version="4.0" xmlns:edmx="http://docs.oasis-open.org/odata/ns/edmx">
  <edmx:DataServices>
    <Schema Namespace="Microsoft.Dynamics.CRM" Alias="mscrm" xmlns="http://docs.oasis-open.org/odata/ns/edm">
      <Action Name="WinOpportunity"><Parameter Name="Status" Type="Edm.Int32" /></Action>
      <Action Name="contoso_DoThing"><Parameter Name="Input" Type="Edm.String" /></Action>
      <Action Name="msdyn_Summarize"><Parameter Name="Input" Type="Edm.String" /></Action>
      <Action Name="new_MyAction"><Parameter Name="Input" Type="Edm.String" /></Action>
    </Schema>
  </edmx:DataServices>
</edmx:Edmx>`;

const MSDYN = '11111111-0000-0000-0000-000000000001';
const CONTOSO = '11111111-0000-0000-0000-000000000002';
const ACTIVE = '11111111-0000-0000-0000-000000000003';
const MESSAGE = '22222222-0000-0000-0000-000000000001';

let solutionsFail = false;

beforeEach(() => {
  getCSDLDocument.mockResolvedValue(CSDL);
  queryData.mockImplementation(async (path: string) => {
    if (path.startsWith('customapis'))
      return {
        value: [
          { uniquename: 'contoso_DoThing', solutionid: CONTOSO },
          { uniquename: 'msdyn_Summarize', solutionid: MSDYN },
        ],
      };
    if (path.startsWith('workflows'))
      // Upper-case GUIDs: the sdkmessage and solution ids still match.
      return {
        value: [{ _sdkmessageid_value: MESSAGE.toUpperCase(), solutionid: CONTOSO.toUpperCase() }],
      };
    if (path.startsWith('sdkmessages'))
      return { value: [{ sdkmessageid: MESSAGE, name: 'new_MyAction' }] };
    if (path.startsWith('solutions')) {
      if (solutionsFail) throw new Error('no read privilege on solution');
      return {
        value: [
          {
            solutionid: MSDYN,
            uniquename: 'msdyn_SalesCopilot',
            publisherid: {
              uniquename: 'microsoftdynamics',
              friendlyname: 'Dynamics 365',
              customizationprefix: 'msdyn',
            },
          },
          {
            solutionid: CONTOSO,
            uniquename: 'ContosoCore',
            publisherid: {
              uniquename: 'contoso',
              friendlyname: 'Contoso',
              customizationprefix: 'contoso',
            },
          },
          {
            solutionid: ACTIVE,
            uniquename: 'Active',
            publisherid: {
              uniquename: 'DefaultPublisherorg1',
              friendlyname: 'Default Publisher for org1',
              customizationprefix: 'new',
            },
          },
        ],
      };
    }
    return { value: [] };
  });
});

const byName = (list: Awaited<ReturnType<typeof fetchCsdlActions>>, name: string) =>
  list.find((a) => a.name === name);

describe('fetchCsdlActions publishers (#248)', () => {
  it('a failed solutions lookup leaves custom messages without a publisher, and is retried a minute later', async () => {
    solutionsFail = true;
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const first = await fetchCsdlActions();
    expect(byName(first, 'contoso_DoThing')).toMatchObject({ source: 'custom-api' });
    expect(byName(first, 'contoso_DoThing')?.publisher).toBeUndefined();
    expect(warn).toHaveBeenCalled();
    warn.mockRestore();

    // A missing privilege won't fix itself: the next load within a minute
    // doesn't re-send the lookup.
    solutionsFail = false;
    const solutionQueries = () =>
      queryData.mock.calls.filter(([path]) => path.startsWith('solutions')).length;
    expect(solutionQueries()).toBe(1);
    expect(byName(await fetchCsdlActions(), 'contoso_DoThing')?.publisher).toBeUndefined();
    expect(solutionQueries()).toBe(1);

    // After a minute it's retried.
    const now = Date.now();
    vi.spyOn(Date, 'now').mockReturnValue(now + 61_000);
    const list = await fetchCsdlActions();
    expect(solutionQueries()).toBe(2);
    expect(byName(list, 'contoso_DoThing')).toMatchObject({
      source: 'custom-api',
      publisher: { name: 'Contoso', microsoft: false },
    });
    expect(byName(list, 'msdyn_Summarize')).toMatchObject({
      source: 'custom-api',
      publisher: { name: 'Dynamics 365', microsoft: true },
    });
    // A custom action: workflow -> sdkmessage -> the workflow's solution.
    // (By its prefix alone it would be the Default Publisher's.)
    expect(byName(list, 'new_MyAction')).toMatchObject({
      source: 'custom-action',
      publisher: { name: 'Contoso', microsoft: false },
    });
    // OOB actions get none.
    expect(byName(list, 'WinOpportunity')?.publisher).toBeUndefined();

    // Complete now, so cached: no more lookups.
    const calls = queryData.mock.calls.length;
    expect(await fetchCsdlActions()).toBe(list);
    expect(queryData.mock.calls.length).toBe(calls);
    vi.restoreAllMocks();
  });
});
