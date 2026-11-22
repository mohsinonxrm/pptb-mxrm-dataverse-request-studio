// #224: with All or nothing on, the Code tab sends Associate / Disassociate's
// requests as one $batch change set in every format that sends them, and
// keeps one call per target otherwise.
import { describe, expect, it, vi } from 'vitest';
import {
  generateCode,
  generatePowerAutomateFields,
  type CodegenInputs,
  type CodeFormat,
} from './codeGenerators';
import { parseBatchRequest } from './batch/multipart';
import type { BuiltRequest } from './urlBuilder';

vi.mock('../env/environment', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../env/environment')>()),
  ENV: { name: 'Test', host: 'contoso.crm.dynamics.com', apiBase: '/api/data/v9.2', user: '' },
}));

const HOST = 'https://contoso.crm.dynamics.com';
const API = '/api/data/v9.2';
const SRC = '00000000-0000-0000-0000-00000000c0de';
const T1 = '00000000-0000-0000-0000-0000000000a1';
const T2 = '00000000-0000-0000-0000-0000000000a2';
const CALLER = '00000000-0000-0000-0000-0000000000c1';
const PATH = `/accounts(${SRC})/contact_customer_accounts/$ref`;

const built: BuiltRequest = {
  relativeUrl: `${API}${PATH}`,
  relativeNoBase: PATH,
  bytes: 0,
  queryParts: [],
  entitySet: 'accounts',
  entityLogical: 'account',
};

function associate(targets: string[], changeSet: boolean): CodegenInputs {
  const reqs = targets.map((t) => ({
    method: 'POST',
    relativeUrl: `${API}${PATH}`,
    body: { '@odata.id': `${HOST}${API}/contacts(${t})` },
    description: `target ${t}`,
  }));
  return {
    method: 'POST',
    built,
    headers: {
      Accept: 'application/json',
      'OData-MaxVersion': '4.0',
      'OData-Version': '4.0',
      'Content-Type': 'application/json; charset=utf-8',
      Authorization: 'Bearer secret',
      MSCRMCallerID: CALLER,
    },
    body: reqs[0].body,
    multiRequests: reqs.length > 1 ? reqs : undefined,
    changeSet,
  };
}

/** The JSON string literals between `start` and `end`, joined with CRLF as the snippet does. */
function bodyFrom(code: string, start: string, end: string): string {
  const from = code.indexOf(start) + start.length;
  const block = code.slice(from, code.indexOf(end, from));
  const lines = block
    .split('\n')
    .map((l) => l.trim().replace(/,$/, ''))
    .filter(Boolean)
    .map((l) => JSON.parse(l) as string);
  return lines.join('\r\n') + '\r\n';
}

function expectChangeSet(body: string) {
  const groups = parseBatchRequest(body, 'multipart/mixed; boundary=batch_drs');
  expect(groups).toHaveLength(1);
  expect(groups[0].atomic).toBe(true);
  expect(groups[0].requests).toEqual(
    [T1, T2].map((t, i) => ({
      method: 'POST',
      url: `${API}${PATH}`,
      // The user's headers on each part; the outer request's own and the
      // credential stay off.
      headers: { mscrmcallerid: CALLER, 'content-type': 'application/json; type=entry' },
      body: JSON.stringify({ '@odata.id': `${HOST}${API}/contacts(${t})` }),
      contentId: String(i + 1),
    })),
  );
}

describe('All or nothing — one change set (#224)', () => {
  const on = associate([T1, T2], true);

  it('fetch posts a $batch whose body is the change set', () => {
    const code = generateCode('fetch', on);
    expect(code).toContain(`await fetch('${HOST}${API}/$batch'`);
    expect(code).toContain("'Content-Type': 'multipart/mixed; boundary=batch_drs'");
    expect(code).not.toContain('secret');
    expectChangeSet(bodyFrom(code, 'const body = [', '].join'));
  });

  it('Python posts the same body', () => {
    const code = generateCode('python', on);
    expect(code).toContain(`requests.post(\n    "${HOST}${API}/$batch"`);
    expectChangeSet(bodyFrom(code, 'body = "\\r\\n".join([', ']) +'));
  });

  it('C# posts the same body', () => {
    const code = generateCode('csharp', on);
    expect(code).toContain('http.PostAsync("api/data/v9.2/$batch", content)');
    expect(code).toContain('MediaTypeHeaderValue.Parse("multipart/mixed; boundary=batch_drs")');
    // C# string escapes match JSON's for this text.
    expectChangeSet(bodyFrom(code, 'new[]\n{', '}) +'));
  });

  it('PowerShell and cURL send it as one POST', () => {
    const ps = generateCode('powershell', on);
    expect(ps).toContain(`-Uri '${HOST}${API}/$batch'`);
    expect(ps).toContain("-ContentType 'multipart/mixed; boundary=batch_drs'");
    expect(ps.match(/Invoke-RestMethod -Method/g)).toHaveLength(1);
    const curl = generateCode('curl', on);
    expect(curl).toContain(`curl -X POST --globoff '${HOST}${API}/$batch'`);
    expect(curl).toContain("printf '%s\\r\\n'");
    expect(curl).toContain('--data-binary @-');
    expect(curl.match(/curl -X/g)).toHaveLength(1);
  });

  it('PowerShell and cURL bodies decode to the same change set', () => {
    // PowerShell '…' doubles a quote; sh '…' writes one as '\''.
    const quoted = (code: string, start: string, end: string, unq: (s: string) => string) => {
      const from = code.indexOf(start) + start.length;
      return (
        code
          .slice(from, code.indexOf(end, from))
          .split('\n')
          .map((l) => l.trim().replace(/ \\$/, ''))
          .filter(Boolean)
          .map((l) => unq(l.slice(1, -1)))
          .join('\r\n') + '\r\n'
      );
    };
    expectChangeSet(
      quoted(generateCode('powershell', on), '$body = (@(', ') -join', (s) =>
        s.replace(/''/g, "'"),
      ),
    );
    expectChangeSet(
      quoted(generateCode('curl', on), "printf '%s\\r\\n' \\", ' |\ncurl', (s) =>
        s.replace(/'\\''/g, "'"),
      ),
    );
  });

  it('Disassociate: one DELETE per part, no body', () => {
    const del: CodegenInputs = {
      method: 'DELETE',
      built,
      headers: { MSCRMCallerID: CALLER },
      multiRequests: [T1, T2].map((t) => ({
        method: 'DELETE',
        relativeUrl: `${API}/accounts(${SRC})/contact_customer_accounts(${t})/$ref`,
      })),
      changeSet: true,
    };
    const code = generateCode('fetch', del);
    const [group] = parseBatchRequest(
      bodyFrom(code, 'const body = [', '].join'),
      'multipart/mixed; boundary=batch_drs',
    );
    expect(group.atomic).toBe(true);
    expect(group.requests).toEqual(
      [T1, T2].map((t, i) => ({
        method: 'DELETE',
        url: `${API}/accounts(${SRC})/contact_customer_accounts(${t})/$ref`,
        headers: { mscrmcallerid: CALLER },
        body: '',
        contentId: String(i + 1),
      })),
    );
  });

  it('xhr and Xrm.WebApi fall back to the fetch change set', () => {
    for (const fmt of ['xhr', 'xrm'] as CodeFormat[]) {
      expect(generateCode(fmt, on)).toContain(`'${HOST}${API}/$batch'`);
    }
  });

  it('executeMultiple nests the requests: a change set', () => {
    const code = generateCode('xrm-batch', on);
    expect(code).toContain('executeMultiple([requests])');
    expect(code).toContain('one change set');
  });

  it('JSON says change-set and keeps every request', () => {
    const out = JSON.parse(generateCode('json', on));
    expect(out.mode).toBe('change-set');
    expect(out.url).toBe(`${HOST}${API}/$batch`);
    expect(out.requests).toHaveLength(2);
  });

  it('Power Automate points at a changeset request', () => {
    const fields = generatePowerAutomateFields(on);
    expect(JSON.stringify(fields)).toContain('Perform a changeset request');
  });
});

describe('All or nothing off, or one target', () => {
  it('keeps one call per target', () => {
    const off = associate([T1, T2], false);
    expect(generateCode('fetch', off)).toContain('for (const r of requests)');
    expect(generateCode('fetch', off)).not.toContain('/$batch');
    expect(JSON.parse(generateCode('json', off)).mode).toBe('multi-request');
    expect(generateCode('xrm-batch', off)).toContain('executeMultiple(requests)');
    expect(JSON.stringify(generatePowerAutomateFields(off))).not.toContain('changeset');
  });

  it('a single target is a single request', () => {
    const one = associate([T1], true);
    for (const fmt of ['fetch', 'python', 'curl', 'csharp', 'powershell'] as CodeFormat[]) {
      expect(generateCode(fmt, one)).not.toContain('/$batch');
    }
  });
});
