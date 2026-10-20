import { describe, expect, it } from 'vitest';
import { csvField, exportFileName, toCsv, type ExportColumn } from './exportResults';

const FV = '@OData.Community.Display.V1.FormattedValue';

const columns: ExportColumn[] = [
  { id: 'name', header: 'Account Name', sourceCol: 'name', showsRaw: false },
  { id: 'statuscode', header: 'Status Reason', sourceCol: 'statuscode', showsRaw: false },
  { id: 'statuscode__raw', header: 'Status Reason (Raw)', sourceCol: 'statuscode', showsRaw: true },
  { id: '_ownerid_value', header: 'Owner', sourceCol: '_ownerid_value', showsRaw: false },
];

const rows = [
  {
    name: 'Contoso, Ltd.',
    statuscode: 1,
    [`statuscode${FV}`]: 'Active',
    _ownerid_value: '0000-guid',
    [`_ownerid_value${FV}`]: 'Jane "JD" Doe',
  },
  { name: 'Line\nbreak', statuscode: null, _ownerid_value: null },
];

describe('csvField', () => {
  it('leaves plain text alone', () => {
    expect(csvField('abc')).toBe('abc');
    expect(csvField('')).toBe('');
  });
  it('quotes commas, quotes, line breaks and edge whitespace', () => {
    expect(csvField('a,b')).toBe('"a,b"');
    expect(csvField('say "hi"')).toBe('"say ""hi"""');
    expect(csvField('a\nb')).toBe('"a\nb"');
    expect(csvField('a\r\nb')).toBe('"a\r\nb"');
    expect(csvField(' lead')).toBe('" lead"');
    expect(csvField('trail ')).toBe('"trail "');
  });
});

describe('toCsv', () => {
  it('exports values as shown, with a BOM and CRLF line breaks', () => {
    const csv = toCsv(rows, columns, 'shown');
    expect(csv.startsWith('﻿')).toBe(true);
    expect(csv.slice(1).split('\r\n')).toEqual([
      'Account Name,Status Reason,Status Reason (Raw),Owner',
      '"Contoso, Ltd.",Active,1,"Jane ""JD"" Doe"',
      '"Line\nbreak",,,',
      '',
    ]);
  });

  it('exports raw values and drops the (Raw) twins', () => {
    const csv = toCsv(rows, columns, 'raw');
    expect(csv.slice(1).split('\r\n')).toEqual([
      'Account Name,Status Reason,Owner',
      '"Contoso, Ltd.",1,0000-guid',
      '"Line\nbreak",,',
      '',
    ]);
  });

  it('serialises object values (collection expands) as JSON', () => {
    const cols: ExportColumn[] = [{ id: 'c', header: 'C', sourceCol: 'c', showsRaw: false }];
    expect(toCsv([{ c: [{ a: 1 }] }], cols, 'shown')).toBe('﻿C\r\n"[{""a"":1}]"\r\n');
  });
});

describe('exportFileName', () => {
  it('stamps the table name with local date and time', () => {
    expect(exportFileName('account', 'csv', new Date(2026, 0, 2, 3, 4, 5))).toBe(
      'account-20260102-030405.csv',
    );
    expect(exportFileName(undefined, 'json', new Date(2026, 9, 17, 13, 0, 0))).toBe(
      'results-20261017-130000.json',
    );
  });
});
