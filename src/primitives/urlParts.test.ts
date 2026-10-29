// #221: ids and names that go into a URL as-is.
import { describe, expect, it } from 'vitest';
import {
  encodeSpaces,
  invalidIdReason,
  invalidNameReason,
  invalidUrlPartReason,
  isODataIdentifier,
} from './urlParts';
import { buildCollection } from '../engine/collectionFormats';

const ID = '00000000-0000-0000-0000-0000000000a1';

describe('invalidIdReason', () => {
  it('passes GUIDs and empty values', () => {
    expect(
      invalidIdReason([
        ['Record id', ID],
        ['Other', null],
        ['Empty', ''],
      ]),
    ).toBeNull();
  });

  it('names the first value that would change the URL', () => {
    expect(invalidIdReason([['Record id', `${ID})/contact_customer_accounts`]])).toBe(
      'Record id "00000000-0000-0000-0000-0000000000a1)/co…" isn\'t a GUID (00000000-0000-0000-0000-000000000000).',
    );
    // Braces are fine to type, not to send.
    expect(invalidIdReason([['Record id', `{${ID}}`]])).toMatch(/isn't a GUID/);
  });
});

describe('invalidNameReason', () => {
  it('accepts identifiers, rejects anything that could add to the path', () => {
    expect(isODataIdentifier('contact_customer_accounts')).toBe(true);
    expect(invalidNameReason([['Relationship', 'parentcustomerid_account']])).toBeNull();
    expect(invalidNameReason([['Relationship', 'nav?$filter=x']])).toMatch(
      /^Relationship "nav\?\$filter=x" isn't a valid name/,
    );
  });

  it('checks ids before names', () => {
    expect(invalidUrlPartReason([['Id', 'x']], [['Name', 'a b']])).toMatch(/^Id/);
  });
});

describe('encodeSpaces', () => {
  it('percent-encodes spaces only', () => {
    expect(encodeSpaces("/accounts?$filter=name eq 'a b'")).toBe(
      "/accounts?$filter=name%20eq%20'a%20b'",
    );
  });

  it('applies to every exported collection format', () => {
    const input = {
      name: 'c',
      baseUrl: 'https://contoso.crm.dynamics.com',
      items: [
        {
          kind: 'request' as const,
          name: 'r',
          method: 'GET',
          path: "/api/data/v9.2/accounts?$filter=name eq 'a'",
          headers: {},
        },
      ],
    };
    for (const fmt of ['postman', 'insomnia', 'bruno'] as const) {
      const out = buildCollection(fmt, input);
      expect(out).toContain('name%20eq%20');
      expect(out).not.toContain("name eq 'a'");
    }
  });
});
