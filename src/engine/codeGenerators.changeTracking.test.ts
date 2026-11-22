// Change tracking in the Code tab (#233): every format that writes the
// Prefer header writes `odata.track-changes` with the rest, from the same
// inputs Retrieve Multiple hands the Code tab. Xrm.WebApi and Power
// Automate's "List rows" can't send it, so they say so and use a request
// that can.
import { beforeAll, describe, expect, it, vi } from 'vitest';

vi.mock('../env/environment', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../env/environment')>()),
  ENV: { name: 'Test', host: 'contoso.crm.dynamics.com', apiBase: '/api/data/v9.2', user: '' },
}));

import { generateCode, generatePowerAutomateFields, type CodeFormat } from './codeGenerators';
import { composeHeaders } from './requestHeaders';
import { buildRetrieveMultiple } from './urlBuilder';
import { emptyPrefer } from '../editors/PreferEditor';
import { registerFixtureTables, retrieveMultipleState } from '../test/fixtures';
import { buildCollection, type CollectionFormat } from './collectionFormats';
import { savedToCollectionItem } from './savedToCollection';
import { serializeRetrieveMultiple, type SavedRequest } from '../state/savedRequests';

beforeAll(() => registerFixtureTables());

const PREFER =
  'odata.include-annotations="OData.Community.Display.V1.FormattedValue", odata.track-changes, odata.maxpagesize=50';

function codeInputs(trackChanges = true) {
  const s = retrieveMultipleState({
    select: ['name'],
    prefer: { ...emptyPrefer(), maxpagesize: 50, trackChanges },
  });
  return { method: 'GET', built: buildRetrieveMultiple(s), headers: composeHeaders(s) };
}

describe('Code tab with Track changes (#233)', () => {
  it('composes one Prefer: annotations, track-changes, then the page size', () => {
    expect(codeInputs().headers.Prefer).toBe(PREFER);
    expect(codeInputs(false).headers.Prefer).not.toContain('track-changes');
  });

  const formats: CodeFormat[] = ['fetch', 'xhr', 'python', 'csharp', 'powershell', 'curl', 'json'];
  it.each(formats)('%s sends Prefer: odata.track-changes', (fmt) => {
    const out = generateCode(fmt, codeInputs());
    expect(out).toContain('odata.track-changes');
    expect(out).toContain('odata.maxpagesize=50');
    expect(generateCode(fmt, codeInputs(false))).not.toContain('odata.track-changes');
  });

  it('fetch, cURL and Python write the whole Prefer header', () => {
    expect(generateCode('fetch', codeInputs())).toContain(`'Prefer': '${PREFER}'`);
    expect(generateCode('curl', codeInputs())).toContain(`-H 'Prefer: ${PREFER}'`);
    expect(generateCode('python', codeInputs())).toContain(`"Prefer": ${JSON.stringify(PREFER)}`);
  });

  it('OpenAPI documents the header', () => {
    expect(generateCode('openapi', codeInputs())).toContain('odata.track-changes');
  });

  it("Xrm.WebApi can't send it, so it says so and writes fetch", () => {
    const out = generateCode('xrm', codeInputs());
    expect(out).toContain(
      "// Xrm.WebApi.retrieveMultipleRecords can't send Prefer: odata.track-changes, so it returns no delta link.",
    );
    expect(out).toContain('odata.track-changes');
    expect(out).not.toContain('Xrm.WebApi.retrieveMultipleRecords(entityLogicalName');
    // Without it, the page size still maps to maxPageSize.
    expect(generateCode('xrm', codeInputs(false))).toContain(
      'retrieveMultipleRecords(entityLogicalName, options, maxPageSize)',
    );
  });

  it('Power Automate uses "Invoke an HTTP request", whose headers can carry it', () => {
    const spec = generatePowerAutomateFields(codeInputs());
    expect(spec.actionName).toBe('Invoke an HTTP request');
    expect(spec.banner).toContain('Prefer: odata.track-changes');
    expect(generateCode('powerautomate', codeInputs())).toContain('"Invoke an HTTP request"');
    expect(generatePowerAutomateFields(codeInputs(false)).actionName).not.toBe(
      'Invoke an HTTP request',
    );
  });

  it('the SDK formats point at RetrieveEntityChangesRequest', () => {
    for (const fmt of ['csharp-sdk', 'powershell-xrm'] as CodeFormat[]) {
      expect(generateCode(fmt, codeInputs())).toContain('RetrieveEntityChangesRequest');
    }
  });

  it.each(['postman', 'insomnia', 'bruno'] as CollectionFormat[])(
    'the %s export of a saved request carries the header',
    (fmt) => {
      const entry = {
        id: 's1',
        name: 'Tracked accounts',
        modeId: 'retrieve-multiple',
        orgScope: 'default',
        savedAt: 1,
        state: serializeRetrieveMultiple(
          retrieveMultipleState({
            select: ['name'],
            prefer: { ...emptyPrefer(), trackChanges: true },
          }),
        ),
      } as SavedRequest;
      const conv = savedToCollectionItem(entry);
      expect(conv.ok).toBe(true);
      if (!conv.ok) return;
      const out = buildCollection(fmt, {
        name: 'C',
        baseUrl: 'https://contoso.crm.dynamics.com',
        items: [conv.item],
      });
      expect(out).toContain('odata.track-changes');
    },
  );
});
