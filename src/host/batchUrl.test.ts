// #225: the URL of a one-part `$batch`. ToolBox takes a path relative to
// /api/data/v9.2/, drops everything after a second literal `?`, and rejects
// `..` and a nested `$batch` with a generic message.
import { describe, expect, it } from 'vitest';
import { BatchPathError, toBatchUrl } from './batchUrl';

const ID = '00000000-0000-0000-0000-0000000000a1';

describe('toBatchUrl', () => {
  it('keeps a relative path as it is', () => {
    expect(toBatchUrl(`accounts(${ID})/name`)).toBe(`accounts(${ID})/name`);
    expect(toBatchUrl("accounts(accountnumber='A1')")).toBe("accounts(accountnumber='A1')");
    expect(toBatchUrl('accounts?$select=name&$top=5')).toBe('accounts?$select=name&$top=5');
  });

  it('strips a leading /', () => {
    expect(toBatchUrl(`/accounts(${ID})`)).toBe(`accounts(${ID})`);
  });

  it('strips an absolute Web API prefix, whatever the v9.x version', () => {
    expect(
      toBatchUrl(`https://contoso.crm.dynamics.com/api/data/v9.2/accounts(${ID})?$select=name`),
    ).toBe(`accounts(${ID})?$select=name`);
    expect(toBatchUrl('https://contoso.crm4.dynamics.com/api/data/v9.0/contacts')).toBe('contacts');
    expect(toBatchUrl('HTTPS://Contoso.crm.dynamics.com/API/DATA/V9.1/contacts')).toBe('contacts');
  });

  it('strips a relative /api/data/v9.x/ prefix', () => {
    expect(toBatchUrl('/api/data/v9.2/accounts')).toBe('accounts');
    expect(toBatchUrl('api/data/v9.2/accounts')).toBe('accounts');
  });

  it('encodes every literal ? after the first, which ToolBox would drop', () => {
    expect(toBatchUrl("accounts?$filter=name eq 'a?b'")).toBe("accounts?$filter=name eq 'a%3Fb'");
    expect(toBatchUrl("accounts?$filter=name eq '??'&$top=1")).toBe(
      "accounts?$filter=name eq '%3F%3F'&$top=1",
    );
    // An absolute nextLink is stripped first, then encoded.
    expect(
      toBatchUrl("https://contoso.crm.dynamics.com/api/data/v9.2/accounts?$filter=name eq 'x?'"),
    ).toBe("accounts?$filter=name eq 'x%3F'");
  });

  it('rejects an empty path', () => {
    expect(() => toBatchUrl('')).toThrow(BatchPathError);
    expect(() => toBatchUrl('   ')).toThrow(/Not sent: the request path is empty/);
    expect(() => toBatchUrl('/')).toThrow(/empty/);
    expect(() => toBatchUrl('https://contoso.crm.dynamics.com/api/data/v9.2/')).toThrow(/empty/);
    expect(() => toBatchUrl('?$top=1')).toThrow(/empty/);
  });

  it('rejects .. and . segments, also percent-encoded', () => {
    expect(() => toBatchUrl('accounts/../systemusers')).toThrow(/can't contain "\.\." segments/);
    expect(() => toBatchUrl('../accounts')).toThrow(BatchPathError);
    expect(() => toBatchUrl('accounts/%2e%2e/systemusers')).toThrow(BatchPathError);
    expect(() => toBatchUrl('accounts/%252e%252e/x')).toThrow(BatchPathError);
    expect(() => toBatchUrl('accounts/./x')).toThrow(BatchPathError);
    // `..` inside a query value isn't a path segment.
    expect(toBatchUrl("accounts?$filter=contains(name,'a..b')")).toBe(
      "accounts?$filter=contains(name,'a..b')",
    );
  });

  it('rejects a % that ToolBox would decode into a broken escape', () => {
    // Live in #225: msfp_emailaddress='z%é…' came back "URL contains invalid encoding".
    const path = "msfp_unsubscribedrecipients(msfp_emailaddress='z%25%C3%A9%40example.invalid')";
    expect(() => toBatchUrl(path)).toThrow(BatchPathError);
    expect(() => toBatchUrl(path)).toThrow(/decodes a "%" in the path/);
    // A trailing % too: `100%25` decodes to `100%`, which isn't an escape.
    expect(() => toBatchUrl("accounts(accountnumber='100%25')")).toThrow(BatchPathError);
  });

  it('keeps a % that decodes cleanly, and any % in the query', () => {
    expect(toBatchUrl("accounts(accountnumber='50%2541')")).toBe(
      "accounts(accountnumber='50%2541')",
    );
    expect(toBatchUrl("accounts?$filter=contains(name,'%25%C3%A9')")).toBe(
      "accounts?$filter=contains(name,'%25%C3%A9')",
    );
  });

  it('rejects a nested $batch', () => {
    expect(() => toBatchUrl('$batch')).toThrow(/inside another \$batch/);
    expect(() => toBatchUrl('/api/data/v9.2/$batch')).toThrow(BatchPathError);
    expect(() => toBatchUrl('%24batch')).toThrow(BatchPathError);
  });

  it("rejects what isn't a Web API path", () => {
    expect(() => toBatchUrl('https://example.com/accounts')).toThrow(/isn't a path under/);
    expect(() => toBatchUrl('//contoso.crm.dynamics.com/accounts')).toThrow(BatchPathError);
    expect(() => toBatchUrl('accounts\\x')).toThrow(BatchPathError);
  });

  it('is a 400 to the executor: nothing was sent', () => {
    try {
      toBatchUrl('');
      expect.unreachable();
    } catch (e) {
      expect((e as BatchPathError).statusCode).toBe(400);
    }
  });
});
