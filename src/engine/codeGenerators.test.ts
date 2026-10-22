import { describe, expect, it, vi } from 'vitest';
import {
  generateCode,
  generatePowerAutomateFields,
  type CodegenInputs,
  type CodeFormat,
} from './codeGenerators';
import type { BuiltRequest } from './urlBuilder';

vi.mock('../mock/environment', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../mock/environment')>()),
  ENV: {
    name: 'Test',
    host: 'contoso.crm.dynamics.com',
    apiBase: '/api/data/v9.2',
    user: '',
  },
}));

const HOST = 'https://contoso.crm.dynamics.com';
const API = '/api/data/v9.2';
const ID = '00000000-0000-0000-0000-0000000000a1';

/** A BuiltRequest for `<API><path>` with the given encoded query parts. */
function built(
  path: string,
  queryParts: { key: string; value: string }[] = [],
  extra: Partial<BuiltRequest> = {},
): BuiltRequest {
  const qs = queryParts.map((p) => `${p.key}=${p.value}`).join('&');
  const noBase = qs ? `${path}?${qs}` : path;
  return {
    relativeUrl: `${API}${noBase}`,
    relativeNoBase: noBase,
    bytes: 0,
    queryParts,
    entitySet: 'accounts',
    entityLogical: 'account',
    ...extra,
  };
}

function inputs(
  method: string,
  b: BuiltRequest,
  extra: Partial<CodegenInputs> = {},
): CodegenInputs {
  return { method, built: b, headers: {}, ...extra };
}

// ── Literal decoders: read a quoted literal back the way its language would ──

/** POSIX sh: a word made of '…' segments and backslash escapes. */
function shUnquote(word: string): string {
  let out = '';
  let i = 0;
  while (i < word.length) {
    const c = word[i];
    if (c === "'") {
      const end = word.indexOf("'", i + 1);
      if (end < 0) throw new Error(`unterminated quote in ${word}`);
      out += word.slice(i + 1, end);
      i = end + 1;
    } else if (c === '\\') {
      out += word[i + 1];
      i += 2;
    } else if (/\s/.test(c)) {
      throw new Error(`unquoted whitespace in ${word}`);
    } else if ('$`"'.includes(c)) {
      throw new Error(`shell would expand / interpret ${c} in ${word}`);
    } else {
      out += c;
      i++;
    }
  }
  return out;
}

/** The single shell word starting at `start` (quoted segments may contain spaces). */
function shWordAt(text: string, start: number): string {
  let i = start;
  let inQuote = false;
  while (i < text.length) {
    const c = text[i];
    if (c === "'") inQuote = !inQuote;
    else if (!inQuote && c === '\\' && i + 1 < text.length && text[i + 1] !== '\n') i++;
    else if (!inQuote && /\s/.test(c)) break;
    i++;
  }
  return text.slice(start, i);
}

/** PowerShell verbatim string: '…' with '' as an escaped quote. */
function psUnquote(lit: string): string {
  expect(lit.startsWith("'") && lit.endsWith("'")).toBe(true);
  const inner = lit.slice(1, -1);
  expect(inner.replace(/''/g, '')).not.toContain("'");
  return inner.replace(/''/g, "'");
}

/** Evaluate a single-quoted JS string literal. */
function jsUnquote(lit: string): string {
  return new Function(`return ${lit};`)() as string;
}

const TRICKY_URL_PATH = '/accounts';
const TRICKY_PARTS = [
  { key: '$select', value: 'name,revenue' },
  { key: '$filter', value: "name eq 'O''Brien $HOME `whoami` \"q\"'" },
];

describe('generateCode — every format', () => {
  it.each<CodeFormat>([
    'fetch',
    'xrm',
    'xrm-batch',
    'xhr',
    'powerautomate',
    'csharp',
    'powershell',
    'curl',
    'json',
  ])('%s produces a snippet for a simple GET', (fmt) => {
    const out = generateCode(fmt, inputs('GET', built('/accounts', [{ key: '$top', value: '5' }])));
    expect(out.trim().length).toBeGreaterThan(0);
  });
});

describe('cURL quoting (#49)', () => {
  it('single-quotes the URL so the shell does not expand $select / $filter', () => {
    const out = generateCode('curl', inputs('GET', built(TRICKY_URL_PATH, TRICKY_PARTS)));
    const start = out.indexOf(`'${HOST}`);
    expect(start).toBeGreaterThan(-1);
    const word = shWordAt(out, start);
    expect(shUnquote(word)).toBe(
      `${HOST}${API}/accounts?$select=name,revenue&$filter=name eq 'O''Brien $HOME \`whoami\` "q"'`,
    );
    expect(out).toContain('--globoff');
  });

  it('quotes header values and the JSON body as literal shell words', () => {
    const out = generateCode(
      'curl',
      inputs('POST', built('/accounts'), {
        headers: { Prefer: 'odata.include-annotations="*"', 'MSCRM.Tag': "it's $x" },
        body: { name: "O'Brien $(rm -rf)" },
      }),
    );
    const headerWords = [...out.matchAll(/-H ('[^\n]*?)(?: \\)?$/gm)].map((m) => m[1]);
    expect(headerWords.map(shUnquote)).toEqual([
      'Prefer: odata.include-annotations="*"',
      "MSCRM.Tag: it's $x",
    ]);
    const bodyStart = out.indexOf('--data-raw ') + '--data-raw '.length;
    expect(JSON.parse(shUnquote(shWordAt(out, bodyStart)))).toEqual({ name: "O'Brien $(rm -rf)" });
  });

  it('never prints a user-supplied Authorization header', () => {
    const out = generateCode(
      'curl',
      inputs('GET', built('/accounts'), { headers: { Authorization: 'Bearer real-secret' } }),
    );
    expect(out).not.toContain('real-secret');
    expect(out).toContain('Bearer <access-token>');
  });
});

describe('PowerShell quoting (#49)', () => {
  it('uses verbatim strings for the URI and header values', () => {
    const out = generateCode(
      'powershell',
      inputs('GET', built(TRICKY_URL_PATH, TRICKY_PARTS), {
        headers: { Prefer: 'odata.include-annotations="*"', "X-It's": "a'b" },
      }),
    );
    const uri = out.match(/-Uri ('(?:[^']|'')*')/)?.[1];
    expect(uri).toBeDefined();
    expect(psUnquote(uri!)).toBe(
      `${HOST}${API}/accounts?$select=name,revenue&$filter=name eq 'O''Brien $HOME \`whoami\` "q"'`,
    );
    const headerLines = [...out.matchAll(/^ {2}('(?:[^']|'')*') = ('(?:[^']|'')*')$/gm)];
    expect(headerLines.map((m) => [psUnquote(m[1]), psUnquote(m[2])])).toEqual([
      ['Prefer', 'odata.include-annotations="*"'],
      ["X-It's", "a'b"],
    ]);
  });

  it('puts the JSON body in a literal here-string', () => {
    const out = generateCode(
      'powershell',
      inputs('PATCH', built(`/accounts(${ID})`, [], { recordId: ID }), {
        body: { name: 'Fabrikam $var' },
      }),
    );
    const m = out.match(/\$body = @'\n([\s\S]*?)\n'@/);
    expect(m).not.toBeNull();
    expect(JSON.parse(m![1])).toEqual({ name: 'Fabrikam $var' });
    expect(out).toContain('-Method PATCH');
  });
});

describe('JavaScript string escaping (#50)', () => {
  it('fetch: the url literal evaluates back to the request URL', () => {
    const out = generateCode('fetch', inputs('GET', built(TRICKY_URL_PATH, TRICKY_PARTS)));
    const lit = out.match(/const url = ('(?:[^'\\\n]|\\.)*');/)?.[1];
    expect(lit).toBeDefined();
    expect(jsUnquote(lit!)).toBe(
      `${HOST}${API}/accounts?$select=name,revenue&$filter=name eq 'O''Brien $HOME \`whoami\` "q"'`,
    );
  });

  it('fetch: a pasted nextLink with a backslash and a trailing newline stays one literal', () => {
    const link = `${HOST}${API}/accounts?$skiptoken=a\\b'c`;
    const out = generateCode(
      'fetch',
      inputs('GET', built('/accounts'), { isNextLink: true, rawNextLink: `${link}\r\n` }),
    );
    const lit = out.match(/const url = ('(?:[^'\\\n]|\\.)*');/)?.[1];
    expect(jsUnquote(lit!)).toBe(link);
  });

  it('fetch: header values with quotes survive as JS literals', () => {
    const out = generateCode(
      'fetch',
      inputs('GET', built('/accounts'), { headers: { 'MSCRM.Tag': "it's \\ fine" } }),
    );
    const m = out.match(/('MSCRM\.Tag'): ('(?:[^'\\\n]|\\.)*')/);
    expect(m).not.toBeNull();
    expect(jsUnquote(m![2])).toBe("it's \\ fine");
  });

  it('Xrm.WebApi: options keep query-significant characters encoded', () => {
    const out = generateCode(
      'xrm',
      inputs(
        'GET',
        built('/accounts', [
          { key: '$select', value: 'name' },
          { key: '$filter', value: "contains(name,'A%26B 50%25 caf%C3%A9 %2B1')" },
        ]),
      ),
    );
    const lit = out.match(/const options = ('(?:[^'\\\n]|\\.)*');/)?.[1];
    // é is decoded for readability; & % + stay encoded so the options
    // string still means the same thing.
    expect(jsUnquote(lit!)).toBe("?$select=name&$filter=contains(name,'A%26B 50%25 café %2B1')");
    expect(out).toContain('Xrm.WebApi.retrieveMultipleRecords(entityLogicalName, options)');
  });

  it('Xrm.WebApi: maps Prefer odata.maxpagesize to the maxPageSize argument', () => {
    const out = generateCode(
      'xrm',
      inputs('GET', built('/accounts'), { headers: { Prefer: 'odata.maxpagesize=50' } }),
    );
    expect(out).toContain('const maxPageSize = 50;');
    expect(out).toContain('retrieveMultipleRecords(entityLogicalName, options, maxPageSize)');
  });

  it('Xrm.WebApi: picks the CRUD helper by verb', () => {
    const b = built(`/accounts(${ID})`, [], { recordId: ID });
    expect(generateCode('xrm', inputs('PATCH', b, { body: { name: 'x' } }))).toContain(
      'Xrm.WebApi.updateRecord(entityLogicalName, id, data)',
    );
    expect(generateCode('xrm', inputs('DELETE', b))).toContain(
      'Xrm.WebApi.deleteRecord(entityLogicalName, id)',
    );
    expect(generateCode('xrm', inputs('GET', b))).toContain('Xrm.WebApi.retrieveRecord(');
  });
});

describe('C# / JSON', () => {
  it('C# doubles quotes inside the verbatim body literal', () => {
    const out = generateCode(
      'csharp',
      inputs('POST', built('/accounts'), { body: { name: 'say "hi"' } }),
    );
    const m = out.match(/new StringContent\(@"((?:[^"]|"")*)"/);
    expect(m).not.toBeNull();
    expect(JSON.parse(m![1].replace(/""/g, '"'))).toEqual({ name: 'say "hi"' });
  });

  it('JSON only carries a body for non-GET requests', () => {
    const get = JSON.parse(
      generateCode('json', inputs('GET', built('/accounts'), { body: { ignored: true } })),
    );
    expect(get).not.toHaveProperty('body');
    expect(get.url).toBe(`${HOST}${API}/accounts`);
    const post = JSON.parse(
      generateCode('json', inputs('POST', built('/accounts'), { body: { a: 1 } })),
    );
    expect(post.body).toEqual({ a: 1 });
  });
});

describe('Power Automate classification (#51)', () => {
  const cases: Array<[string, string, BuiltRequest, Record<string, unknown> | undefined, string]> =
    [
      [
        'GET collection',
        'GET',
        built('/accounts', [{ key: '$select', value: 'name' }]),
        undefined,
        'List rows',
      ],
      [
        'GET row',
        'GET',
        built(`/accounts(${ID})`, [], { recordId: ID }),
        undefined,
        'Get a row by ID',
      ],
      ['POST collection', 'POST', built('/accounts'), { name: 'x' }, 'Add a new row'],
      [
        'PATCH row',
        'PATCH',
        built(`/accounts(${ID})`, [], { recordId: ID }),
        { name: 'x' },
        'Update a row',
      ],
      [
        'DELETE row',
        'DELETE',
        built(`/accounts(${ID})`, [], { recordId: ID }),
        undefined,
        'Delete a row',
      ],
      ['DELETE placeholder row', 'DELETE', built('/accounts(<id>)'), undefined, 'Delete a row'],
      [
        'bound action',
        'POST',
        built(`/accounts(${ID})/Microsoft.Dynamics.CRM.sample_DoThing`, [], { recordId: ID }),
        { Input: 1 },
        'Perform a bound action',
      ],
      [
        'unbound action',
        'POST',
        built('/WinOpportunity', [], { entitySet: '', entityLogical: '' }),
        { Status: 3 },
        'Perform an unbound action',
      ],
      [
        'function',
        'GET',
        built('/WhoAmI()', [], { entitySet: '', entityLogical: '' }),
        undefined,
        'Invoke an HTTP request (function)',
      ],
      [
        'DELETE column',
        'DELETE',
        built(`/accounts(${ID})/revenue`, [], { recordId: ID }),
        undefined,
        'Invoke an HTTP request',
      ],
      [
        'PUT column',
        'PUT',
        built(`/accounts(${ID})/name`, [], { recordId: ID }),
        { value: 'x' },
        'Invoke an HTTP request',
      ],
      [
        'POST $ref',
        'POST',
        built(`/accounts(${ID})/contact_customer_accounts/$ref`, [], { recordId: ID }),
        { '@odata.id': 'x' },
        'Invoke an HTTP request',
      ],
      [
        'PATCH by alternate key',
        'PATCH',
        built("/accounts(accountnumber='A1')"),
        { name: 'x' },
        'Invoke an HTTP request',
      ],
    ];

  it.each(cases)('%s → %s', (_label, method, b, body, expected) => {
    const spec = generatePowerAutomateFields(inputs(method, b, { body }));
    expect(spec.actionName).toBe(expected);
  });

  it('warns that a column DELETE does not delete the row, in both renderings', () => {
    const i = inputs('DELETE', built(`/accounts(${ID})/revenue`, [], { recordId: ID }));
    const spec = generatePowerAutomateFields(i);
    expect(spec.banner).toMatch(/does NOT delete the row/);
    expect(spec.fields.find((f) => f.label === 'Relative URI')?.value).toBe(
      `/accounts(${ID})/revenue`,
    );

    const text = generateCode('powerautomate', i);
    expect(text).toMatch(/⟶\s+"Invoke an HTTP request" action/);
    expect(text).not.toMatch(/⟶\s+"Delete a row"/);
    expect(text).toContain('does NOT delete the row');
  });

  it('decodes List rows field values and maps Prefer to Maximum Pagination Size', () => {
    const spec = generatePowerAutomateFields(
      inputs(
        'GET',
        built('/accounts', [
          { key: '$select', value: 'name,_primarycontactid_value' },
          { key: '$filter', value: "contains(name,'A%26B')" },
          { key: '$count', value: 'true' },
        ]),
        { headers: { Prefer: 'odata.maxpagesize=100' } },
      ),
    );
    const field = (label: string) => spec.fields.find((f) => f.label === label)?.value;
    expect(field('Table name')).toBe('accounts');
    expect(field('Select Columns')).toBe('name,_primarycontactid_value');
    expect(field('Filter Rows')).toBe("contains(name,'A&B')");
    expect(field('Row Count')).toBe('Yes');
    expect(field('Maximum Pagination Size')).toBe('100');
    expect(spec.httpFallback).toBeUndefined();
  });

  it('routes $apply to the HTTP fallback and saved queries to the predefined form', () => {
    const apply = generatePowerAutomateFields(
      inputs('GET', built('/accounts', [{ key: '$apply', value: 'aggregate($count as n)' }])),
    );
    expect(apply.actionName).toBe('List rows');
    expect(apply.httpFallback?.uri).toBe(`${HOST}${API}/accounts?$apply=aggregate($count as n)`);

    const saved = generatePowerAutomateFields(
      inputs('GET', built('/accounts', [{ key: 'savedQuery', value: ID }])),
    );
    expect(saved.actionName).toMatch(/Saved \/ User Query/);
    expect(saved.fields.find((f) => f.label === 'Saved Query')?.value).toBe(ID);
  });

  it('splits lookup binds into GUID fields for Add a new row', () => {
    const spec = generatePowerAutomateFields(
      inputs('POST', built('/accounts'), {
        body: { name: 'x', 'primarycontactid@odata.bind': `/contacts(${ID})` },
      }),
    );
    expect(spec.fields.find((f) => f.label === 'primarycontactid')).toMatchObject({
      value: ID,
      hint: expect.stringContaining('contacts'),
    });
  });
});

describe('C# header strings (#244, CodeQL js/incomplete-sanitization)', () => {
  it('escapes backslashes, quotes and line breaks in header values', () => {
    const code = generateCode(
      'csharp',
      inputs('GET', built('/accounts'), {
        // Header values C:\temp\x"y and a<LF>b.
        headers: { 'X-Path': String.raw`C:\temp\x"y`, 'X-Lines': 'a\nb' },
      }),
    );
    // C# literals "C:\\temp\\x\"y" and "a\nb".
    expect(code).toContain(String.raw`TryAddWithoutValidation("X-Path", "C:\\temp\\x\"y");`);
    expect(code).toContain(String.raw`TryAddWithoutValidation("X-Lines", "a\nb");`);
  });
});
