// Reading pasted curl commands (#271): bash, cmd and PowerShell quoting, and
// the options that matter.
import { describe, expect, it } from 'vitest';
import { parseCurl, shellWords } from './curlCommand';

const ok = (text: string) => {
  const r = parseCurl(text);
  if (!r.ok) throw new Error(r.error);
  return r.command;
};
const err = (text: string) => {
  const r = parseCurl(text);
  if (r.ok) throw new Error('expected an error');
  return r.error;
};

const URL = 'https://contoso.crm.dynamics.com/api/data/v9.2/accounts';

describe('shellWords', () => {
  it('keeps $ and backticks literal and reads bash quoting', () => {
    expect(shellWords(`curl '${URL}?$select=name' "a \\"b\\" $HOME \`x\`" c\\ d`)).toEqual([
      'curl',
      `${URL}?$select=name`,
      'a "b" $HOME `x`',
      'c d',
    ]);
  });

  it("reads $'…' with escapes, as Chrome's Copy as cURL (bash) writes bodies", () => {
    expect(shellWords(`--data-raw $'{"name":"O\\'Brien\\\\x\\n\\u00e9"}'`)).toEqual([
      '--data-raw',
      `{"name":"O'Brien\\x\né"}`,
    ]);
  });

  it('joins continued lines (bash \\, PowerShell `, cmd ^)', () => {
    expect(shellWords(`curl \\\n  -H 'A: 1' \\\r\n  '${URL}'`)).toEqual([
      'curl',
      '-H',
      'A: 1',
      URL,
    ]);
    expect(shellWords(`curl.exe \`\n  -H "A: 1"`)).toEqual(['curl.exe', '-H', 'A: 1']);
    expect(
      shellWords(
        `curl ^"${URL}^" ^\n  -H ^"content-type: application/json^" ^\n  --data-raw ^"^{^\\^"name^\\^":^\\^"A ^& B^\\^"^}^"`,
      ),
    ).toEqual([
      'curl',
      URL,
      '-H',
      'content-type: application/json',
      '--data-raw',
      '{"name":"A & B"}',
    ]);
  });

  it("skips comments, as at the top of DRS's cURL code", () => {
    expect(shellWords("# it's a comment\ncurl 'a#b' x # trailing\n")).toEqual(['curl', 'a#b', 'x']);
  });

  it('stops at a pipe or a second command', () => {
    expect(shellWords(`curl '${URL}' | jq . ; echo done`)).toEqual(['curl', URL]);
  });

  it('refuses an unclosed quote', () => {
    expect(() => shellWords(`curl 'oops`)).toThrow(/never closed/);
  });
});

describe('parseCurl', () => {
  it('reads method, URL, headers and body', () => {
    expect(
      ok(
        `curl -X PATCH '${URL}(1)' -H 'If-Match: *' -H 'Content-Type: application/json' --data-raw '{"name":"A"}'`,
      ),
    ).toEqual({
      method: 'PATCH',
      url: `${URL}(1)`,
      headers: [
        { name: 'If-Match', value: '*' },
        { name: 'Content-Type', value: 'application/json' },
      ],
      body: '{"name":"A"}',
      notes: [],
    });
  });

  it('defaults to GET, and to POST when there is a body', () => {
    expect(ok(`curl ${URL}`).method).toBe('GET');
    expect(ok(`curl ${URL} -d '{}'`).method).toBe('POST');
    expect(ok(`curl --request delete --url ${URL}`).method).toBe('DELETE');
  });

  it('reads bundled and attached short options', () => {
    const c = ok(`curl -sSL -XPOST -H'Accept: application/json' ${URL} -d'{}'`);
    expect(c.method).toBe('POST');
    expect(c.headers).toEqual([{ name: 'Accept', value: 'application/json' }]);
    expect(c.body).toBe('{}');
  });

  it('reads --name=value', () => {
    expect(ok(`curl --request=DELETE --url=${URL}`)).toMatchObject({ method: 'DELETE', url: URL });
  });

  it('joins repeated -d with &, and -G moves the data to the query', () => {
    expect(ok(`curl ${URL} -d a=1 -d b=2`).body).toBe('a=1&b=2');
    const g = ok(`curl -G ${URL}?x=1 --data-urlencode '$top=5'`);
    expect(g).toMatchObject({ method: 'GET', url: `${URL}?x=1&$top=5`, body: undefined });
  });

  it('--json sets the JSON headers and POST', () => {
    expect(ok(`curl --json '{"a":1}' ${URL}`)).toMatchObject({
      method: 'POST',
      headers: [
        { name: 'Content-Type', value: 'application/json' },
        { name: 'Accept', value: 'application/json' },
      ],
    });
  });

  it('notes credentials and ignored options instead of keeping them', () => {
    const c = ok(`curl -u me:secret -b 'a=b' --frobnicate ${URL}`);
    expect(c.notes.join(' ')).toMatch(/Left out -u.*Left out -b.*Ignored --frobnicate/);
    expect(JSON.stringify(c)).not.toContain('secret');
  });

  it("refuses what it can't read", () => {
    expect(err(`wget ${URL}`)).toMatch(/starts with curl/);
    expect(err('curl -H "A: 1"')).toMatch(/no URL/);
    expect(err(`curl ${URL} -d @body.json`)).toMatch(/read from a file \(@body\.json\)/);
    expect(err(`curl ${URL} -F file=@a.png`)).toMatch(/uploads a file or form/);
    expect(err(`curl -I ${URL}`)).toMatch(/HEAD/);
    expect(err(`curl ${URL} ${URL}`)).toMatch(/more than one URL/);
    expect(err(`curl ${URL} -H`)).toMatch(/-H needs a value/);
  });

  it('keeps --data-raw @ literal', () => {
    expect(ok(`curl ${URL} --data-raw '@odata'`).body).toBe('@odata');
  });

  it('accepts curl.exe and a leading $ prompt', () => {
    expect(ok(`$ curl.exe ${URL}`).url).toBe(URL);
  });
});

describe('parseCurl, from the #271 review', () => {
  it('reads bundled flags that end in an option with a value', () => {
    expect(ok(`curl -sX DELETE '${URL}(1)'`)).toMatchObject({ method: 'DELETE', url: `${URL}(1)` });
    expect(ok(`curl '${URL}' -sd '{"a":1}'`)).toMatchObject({ method: 'POST', body: '{"a":1}' });
    expect(ok(`curl -sSXPATCH '${URL}(1)'`).method).toBe('PATCH');
  });

  it('asks for one command at a time', () => {
    expect(err(`curl '${URL}'\ncurl '${URL}'`)).toMatch(/one curl command at a time/);
  });

  it("doesn't repeat a header value that has no name: it may be a token", () => {
    const c = ok(`curl '${URL}' -H 'Authorization Bearer secret'`);
    expect(c.notes.join(' ')).toMatch(/Left out the header "Authorization"/);
    expect(JSON.stringify(c)).not.toContain('secret');
  });
});
