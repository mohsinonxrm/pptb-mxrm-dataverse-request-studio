// suppliedColumnTest: the "will this column be sent with a value" rule behind
// the field set's required count, "N populated" badge and missing marker (#210).
import { describe, expect, it } from 'vitest';
import { suppliedColumnTest } from './fieldBody';
import { account } from '../test/fixtures';

const col = (name: string) => account.columns.find((c) => c.logicalName === name)!;

describe('suppliedColumnTest (Form mode)', () => {
  it('ignores seeded empty values', () => {
    const supplied = suppliedColumnTest(account, null, {
      name: '',
      sample_tags: [],
      primarycontactid: { id: '' } as never,
      industrycode: null as never,
    });
    expect(
      ['name', 'sample_tags', 'primarycontactid', 'industrycode'].filter((c) => supplied(col(c))),
    ).toEqual([]);
  });

  it('counts real values, including false and 0', () => {
    const supplied = suppliedColumnTest(account, null, {
      name: 'Contoso',
      donotemail: false,
      numberofemployees: 0,
    });
    expect(['name', 'donotemail', 'numberofemployees'].every((c) => supplied(col(c)))).toBe(true);
    expect(supplied(col('revenue'))).toBe(false);
  });
});

describe('suppliedColumnTest (JSON mode)', () => {
  it('reads the typed body, not the form values', () => {
    const supplied = suppliedColumnTest(account, '{"name": "Contoso"}', { accountnumber: 'A-1' });
    expect(supplied(col('name'))).toBe(true);
    expect(supplied(col('accountnumber'))).toBe(false);
  });
});
