// Python (requests) code generation (#270): golden snippets per mode group,
// and string literals that Python reads back as the exact value.
import { describe, expect, it, vi } from 'vitest';
import { generateCode, pyStr, type CodegenInputs } from './codeGenerators';
import type { BuiltRequest } from './urlBuilder';
import { compileBatch, newGroup, newItem, planDispatch } from './batch/batchModel';
import { generateBatchCode, toEncodableGroups } from './batch/batchCode';
import { encodeBatch } from './batch/multipart';

vi.mock('../env/environment', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../env/environment')>()),
  ENV: {
    name: 'Test',
    host: 'contoso.crm.dynamics.com',
    apiBase: '/api/data/v9.2',
    user: '',
  },
}));

const API = '/api/data/v9.2';
const ID = '00000000-0000-0000-0000-0000000000a1';

function built(path: string, query = '', extra: Partial<BuiltRequest> = {}): BuiltRequest {
  const noBase = query ? `${path}?${query}` : path;
  return {
    relativeUrl: `${API}${noBase}`,
    relativeNoBase: noBase,
    bytes: 0,
    queryParts: [],
    entitySet: 'accounts',
    entityLogical: 'account',
    ...extra,
  };
}

const py = (method: string, b: BuiltRequest, extra: Partial<CodegenInputs> = {}) =>
  generateCode('python', { method, built: b, headers: {}, ...extra });

/**
 * Read a double-quoted Python string literal the way Python does — only the
 * escapes the generator can emit; anything else (a raw line break, an
 * unknown escape) fails the test.
 */
function pyUnquote(lit: string): string {
  expect(lit.startsWith('"') && lit.endsWith('"')).toBe(true);
  const SIMPLE: Record<string, string> = {
    '\\': '\\',
    '"': '"',
    n: '\n',
    r: '\r',
    t: '\t',
    b: '\b',
    f: '\f',
  };
  let out = '';
  for (let i = 1; i < lit.length - 1; i++) {
    const c = lit[i];
    if (c === '"' || c === '\n' || c === '\r') throw new Error(`bare ${JSON.stringify(c)}`);
    if (c !== '\\') {
      out += c;
      continue;
    }
    const e = lit[++i];
    if (e in SIMPLE) out += SIMPLE[e];
    else if (e === 'u') {
      out += String.fromCharCode(parseInt(lit.slice(i + 1, i + 5), 16));
      i += 4;
    } else throw new Error(`unexpected escape \\${e}`);
  }
  return out;
}

describe('Python — reads', () => {
  it('a Retrieve Multiple GET', () => {
    expect(
      py('GET', built('/accounts', "$select=name&$filter=name eq 'A'&$top=5"), {
        headers: { Prefer: 'odata.include-annotations="*"', 'OData-Version': '4.0' },
      }),
    ).toBe(`# Dataverse Web API — Python (requests)
# pip install requests. Replace <access-token> with a bearer token for
# https://contoso.crm.dynamics.com (msal or azure-identity can get one).

import requests

ACCESS_TOKEN = "<access-token>"

url = "https://contoso.crm.dynamics.com/api/data/v9.2/accounts?$select=name&$filter=name%20eq%20'A'&$top=5"
headers = {
    "Authorization": "Bearer " + ACCESS_TOKEN,
    "Prefer": "odata.include-annotations=\\"*\\"",
    "OData-Version": "4.0",
}

res = requests.request(
    "GET",
    url,
    headers=headers,
    timeout=60,
)

if not res.ok:
    try:
        message = res.json()["error"]["message"]
    except (ValueError, KeyError, TypeError):
        message = res.reason
    raise SystemExit(f"Dataverse {res.status_code}: {message}")

# 204 No Content (e.g. Delete) doesn't carry a body
if res.status_code == 204 or not res.content:
    print("Status:", res.status_code)
elif "json" in res.headers.get("Content-Type", ""):
    data = res.json()
    print(data)
    if isinstance(data, dict) and isinstance(data.get("value"), list):
        print(f"Returned {len(data['value'])} record(s)")
        if "@odata.count" in data:
            print("@odata.count:", data["@odata.count"])
        if "@odata.nextLink" in data:
            print("nextLink:", data["@odata.nextLink"])
elif res.headers.get("Content-Type", "").startswith("text/"):
    print(res.text)
else:
    print(f"{len(res.content)} bytes ({res.headers.get('Content-Type')})")
`);
  });

  it('a function call keeps its parameter aliases in the URL', () => {
    const code = py(
      'GET',
      built('/RetrieveTotalRecordCount(EntityNames=@p1)', "@p1=%5B'account'%2C'contact'%5D", {
        entitySet: undefined,
        entityLogical: undefined,
      }),
    );
    expect(code).toContain(
      `url = "https://contoso.crm.dynamics.com${API}/RetrieveTotalRecordCount(EntityNames=@p1)?@p1=%5B'account'%2C'contact'%5D"\n`,
    );
    expect(code).toContain('    "GET",\n    url,\n    headers=headers,\n    timeout=60,\n)');
  });

  it('a nextLink is sent as pasted, trimmed', () => {
    const code = py('GET', built('/accounts'), {
      isNextLink: true,
      rawNextLink:
        '  https://contoso.crm.dynamics.com/api/data/v9.2/accounts?$skiptoken=%3Ccookie%3E\n',
    });
    expect(code).toContain(
      'url = "https://contoso.crm.dynamics.com/api/data/v9.2/accounts?$skiptoken=%3Ccookie%3E"\n',
    );
  });
});

describe('Python — writes', () => {
  it('a Create sends the body as a dict through json=', () => {
    const code = py('POST', built('/accounts'), {
      headers: { 'Content-Type': 'application/json', Prefer: 'return=representation' },
      body: {
        name: 'Contoso',
        revenue: 1000.5,
        creditonhold: false,
        description: null,
        'primarycontactid@odata.bind': `/contacts(${ID})`,
        contact_customer_accounts: [{ firstname: 'Ann', donotemail: true }],
        empty: {},
      },
    });
    expect(code).toContain(`payload = {
    "name": "Contoso",
    "revenue": 1000.5,
    "creditonhold": False,
    "description": None,
    "primarycontactid@odata.bind": "/contacts(${ID})",
    "contact_customer_accounts": [
        {
            "firstname": "Ann",
            "donotemail": True,
        },
    ],
    "empty": {},
}

res = requests.request(
    "POST",
    url,
    headers=headers,
    json=payload,
    timeout=60,
)`);
    expect(code).toContain('    "Content-Type": "application/json",\n');
  });

  it('numbers keep the value JSON would send', () => {
    const code = py('POST', built('/accounts'), {
      body: { big: Number('12345678901234567890'), e: 1e21, tiny: 1e-7, neg: -0, dec: -5.25 },
    });
    expect(code).toContain(`payload = {
    "big": 12345678901234567000,
    "e": 1e+21,
    "tiny": 1e-7,
    "neg": 0,
    "dec": -5.25,
}`);
  });

  it('an Update with no body still sends an empty object', () => {
    expect(py('PATCH', built(`/accounts(${ID})`))).toContain(
      'payload = {}\n\nres = requests.request(\n    "PATCH",',
    );
  });

  it('a Delete has no body', () => {
    const code = py('DELETE', built(`/accounts(${ID})`), { headers: { 'If-Match': '*' } });
    expect(code).toContain('    "DELETE",\n    url,\n    headers=headers,\n    timeout=60,\n)');
    expect(code).not.toContain('payload');
    expect(code).not.toContain('json=');
  });

  it('a file upload reads the bytes from a file', () => {
    const code = py('PATCH', built(`/accounts(${ID})/entityimage`), {
      rawBody: '<file bytes>',
      headers: { 'Content-Type': 'application/octet-stream' },
    });
    expect(code).toContain(
      '# <file bytes>: the file content.\nwith open("<path-to-file>", "rb") as f:\n    file_bytes = f.read()\n',
    );
    expect(code).toContain('    data=file_bytes,\n');
    expect(code).not.toContain('json=');
  });
});

describe('Python — multi-request', () => {
  it('sends each Associate in order on one session', () => {
    const code = py('POST', built(`/accounts(${ID})/contact_customer_accounts/$ref`), {
      headers: { 'OData-Version': '4.0', Authorization: 'Bearer secret' },
      multiRequests: [1, 2].map((n) => ({
        method: 'POST',
        relativeUrl: `${API}/accounts(${ID})/contact_customer_accounts/$ref`,
        body: { '@odata.id': `https://contoso.crm.dynamics.com${API}/contacts(${n})` },
        description: `Contact ${n}`,
      })),
    });
    expect(code).toContain(`calls = [
    # Contact 1
    {"method": "POST", "url": "https://contoso.crm.dynamics.com/api/data/v9.2/accounts(${ID})/contact_customer_accounts/$ref", "json": {
        "@odata.id": "https://contoso.crm.dynamics.com/api/data/v9.2/contacts(1)",
    }},
    # Contact 2`);
    expect(code).toContain(`session.headers.update({
    "Authorization": "Bearer " + ACCESS_TOKEN,
    "OData-Version": "4.0",
})`);
    expect(code).not.toContain('secret');
  });

  it('keeps per-request headers and placeholder URLs of a chunked upload', () => {
    const code = py('PATCH', built(`/accounts(${ID})/filecolumn`), {
      multiRequests: [
        {
          method: 'PATCH',
          relativeUrl: `${API}/accounts(${ID})/filecolumn`,
          headers: { 'x-ms-transfer-mode': 'chunked' },
          description: 'Start the upload\nthen send chunks',
        },
        {
          method: 'PATCH',
          relativeUrl: '<upload-url-from-Location-header>',
          headers: { 'Content-Range': 'bytes 0-4194303/*' },
          rawBody: '<chunk bytes>',
        },
      ],
    });
    expect(code).toContain('    # Start the upload then send chunks\n');
    expect(code).toContain(
      '{"method": "PATCH", "url": "<upload-url-from-Location-header>", "headers": {"Content-Range": "bytes 0-4194303/*"}, "data": "<chunk bytes>"},  # replace with the bytes',
    );
    expect(code).toContain('"headers": {"x-ms-transfer-mode": "chunked"}},');
  });
});

describe('Python string literals', () => {
  const TRICKY = [
    'plain',
    'O\'Brien "q" C:\\temp\\x',
    'line1\nline2\r\n\ttab',
    '{name} %s $HOME `x` \\N{BULLET} \\u0041',
    'ünïcødé ✓ 😀 \u2028 \u2029 \u0000 \u001f \u007f',
  ];

  it.each(TRICKY)('pyStr round-trips %j', (s) => {
    expect(pyUnquote(pyStr(s))).toBe(s);
  });

  it('header names and values, the URL and body strings read back exactly', () => {
    const value = TRICKY[3] + TRICKY[1] + TRICKY[2];
    const code = py('POST', built('/accounts', "$filter=name eq 'a\"b'"), {
      headers: { 'X-Odd': value },
      body: { [value]: value },
    });
    const lit = (prefix: string) => {
      const line = code.split('\n').find((l) => l.startsWith(prefix))!;
      return line.slice(prefix.length).replace(/,$/, '');
    };
    expect(pyUnquote(lit('    "X-Odd": '))).toBe(value);
    expect(pyUnquote(lit('url = '))).toBe(
      `https://contoso.crm.dynamics.com${API}/accounts?$filter=name%20eq%20'a"b'`,
    );
    const bodyLine = code.split('\n').find((l) => l.startsWith('    "{name}'))!;
    const [k, v] = bodyLine.trim().replace(/,$/, '').split('": "');
    expect(pyUnquote(`${k}"`)).toBe(value);
    expect(pyUnquote(`"${v}`)).toBe(value);
  });

  it('a line break in a comment would end it, and a NUL is refused — both are folded', () => {
    const code = py('PATCH', built(`/accounts(${ID})/entityimage`), { rawBody: 'a\nb\u0000c' });
    expect(code).toContain('# a b c: the file content.');
    // eslint-disable-next-line no-control-regex
    expect(code).not.toMatch(/[\u0000-\u0008\u000b-\u001f]/);
  });

  it('a lone surrogate in a comment becomes U+FFFD; a pair is kept', () => {
    const code = py('PATCH', built(`/accounts(${ID})/entityimage`), {
      rawBody: 'a\ud800b\udc00c😀',
    });
    expect(code).toContain('# a�b�c😀: the file content.');
  });
});

describe('Python — $batch', () => {
  it('posts the multipart body joined with CRLF', () => {
    const { items } = compileBatch([
      newGroup(false, [newItem('create', { entitySet: 'accounts', body: '{"name":"A \\"q\\""}' })]),
    ]);
    const code = generateBatchCode('python', {
      items,
      plan: planDispatch(items),
      continueOnError: true,
      orgUrl: 'https://contoso.crm.dynamics.com',
      boundaries: { batch: 'batch_1', changeset: (i) => `changeset_${i}` },
    });
    expect(code).toContain('body = "\\r\\n".join([\n    "--batch_1",\n');
    expect(code).toContain('    "  \\"name\\": \\"A \\\\\\"q\\\\\\"\\"",\n');
    expect(code).toContain(']) + "\\r\\n"\n');
    expect(code).toContain('    "https://contoso.crm.dynamics.com/api/data/v9.2/$batch",\n');
    expect(code).toContain('        "Content-Type": "multipart/mixed; boundary=batch_1",\n');
    expect(code).toContain('        "Prefer": "odata.continue-on-error",\n');
    expect(code).toContain('    data=body.encode("utf-8"),\n');
    expect(code).not.toContain('If-None-Match');
  });

  it('the joined lines are the exact multipart body, change sets included', () => {
    const groups = [
      newGroup(true, [
        newItem('create', { entitySet: 'accounts', body: '{"name":"ü \\\\ \\"q\\""}' }),
        newItem('update', { entitySet: 'accounts', recordId: '1', body: '{"name":"B"}' }),
      ]),
      newGroup(false, [newItem('delete', { entitySet: 'accounts', recordId: '2' })]),
    ];
    const { items } = compileBatch(groups);
    const boundaries = { batch: 'batch_1', changeset: (i: number) => `changeset_${i}` };
    const code = generateBatchCode('python', {
      items,
      plan: planDispatch(items),
      continueOnError: false,
      orgUrl: '',
      boundaries,
    });
    const listed = code.slice(
      code.indexOf('".join([\n') + '".join([\n'.length,
      code.indexOf('\n]) + "\\r\\n"'),
    );
    const sent = listed.split('\n').map((l) => pyUnquote(l.trim().replace(/,$/, '')));
    expect(sent.join('\r\n') + '\r\n').toBe(encodeBatch(toEncodableGroups(items), boundaries).body);
  });
});
