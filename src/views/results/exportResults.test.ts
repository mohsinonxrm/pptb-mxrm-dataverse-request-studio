import { describe, expect, it } from 'vitest';
import {
  csvField,
  exportFileName,
  spreadsheetSafe,
  tsvCell,
  toCsv,
  type ExportColumn,
} from './exportResults';

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

describe('spreadsheetSafe (#239)', () => {
  it('prefixes text that a spreadsheet would run as a formula', () => {
    expect(spreadsheetSafe('=1+2')).toBe("'=1+2");
    expect(spreadsheetSafe('+1+2')).toBe("'+1+2");
    expect(spreadsheetSafe('-1+2')).toBe("'-1+2");
    expect(spreadsheetSafe('@SUM(1,1)')).toBe("'@SUM(1,1)");
    expect(spreadsheetSafe('=HYPERLINK("https://example.invalid","x")')).toBe(
      '\'=HYPERLINK("https://example.invalid","x")',
    );
    expect(spreadsheetSafe('\t=1')).toBe("'\t=1");
    expect(spreadsheetSafe('\r=1')).toBe("'\r=1");
    expect(spreadsheetSafe('  =1+2')).toBe("'  =1+2");
  });

  it('leaves ordinary text and plain numbers alone', () => {
    for (const t of [
      '',
      'Contoso',
      'a=b',
      'x-1',
      '2026-10-20',
      '1+2',
      '-5',
      '+1.5e3',
      '-.5',
      ' -12 ',
    ])
      expect(spreadsheetSafe(t)).toBe(t);
  });

  it("keeps a number's formatted sign but still catches = and @", () => {
    expect(spreadsheetSafe('-$1,234.00', -1234)).toBe('-$1,234.00');
    expect(spreadsheetSafe('-1.234,50 €', -1234.5)).toBe('-1.234,50 €');
    expect(spreadsheetSafe('+5', 5)).toBe('+5');
    // A choice label (number source) that starts like a formula.
    expect(spreadsheetSafe('=cmd', 1)).toBe("'=cmd");
    expect(spreadsheetSafe('@x', true)).toBe("'@x");
    // Text that only looks numeric-sourced is still checked as text.
    expect(spreadsheetSafe('-1+2', '-1+2')).toBe("'-1+2");
  });
});

describe('toCsv formula neutralising (#239)', () => {
  const cols: ExportColumn[] = [
    { id: 'name', header: 'Name', sourceCol: 'name', showsRaw: false },
    { id: 'amt', header: '=Amount', sourceCol: 'amt', showsRaw: false },
  ];
  const data = [
    { name: '=1+2', amt: -5, [`amt${FV}`]: '-$5.00' },
    { name: '@SUM(1,1), more', amt: 3 },
    { name: '-cmd|calc', amt: null },
  ];

  it('neutralises cells and headers, then quotes as usual', () => {
    expect(toCsv(data, cols, 'shown').slice(1).split('\r\n')).toEqual([
      "Name,'=Amount",
      "'=1+2,-$5.00",
      '"\'@SUM(1,1), more",3',
      "'-cmd|calc,",
      '',
    ]);
  });

  it('exports raw negative numbers unchanged', () => {
    expect(toCsv(data, cols, 'raw').slice(1).split('\r\n')[1]).toBe("'=1+2,-5");
  });
});

describe('tsvCell (#239 review)', () => {
  it('also neutralises a leading double quote, and flattens tabs and line breaks', () => {
    expect(tsvCell('"=HYPERLINK(""x"")"')).toBe(`'"=HYPERLINK(""x"")"`);
    expect(tsvCell('  "quoted"')).toBe(`'  "quoted"`);
    expect(tsvCell('=1+2')).toBe("'=1+2");
    expect(tsvCell('a\tb\nc')).toBe('a b c');
    expect(tsvCell('-$5.00', -5)).toBe('-$5.00');
    expect(tsvCell('say "hi"')).toBe('say "hi"');
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
