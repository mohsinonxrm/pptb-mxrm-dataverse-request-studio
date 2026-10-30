// DateTime column values: DateOnly is a calendar date in every zone (#92),
// TimeZoneIndependent keeps the digits as typed, and UserLocal converts the
// local time to UTC (#106). Each case pins the browser's time zone.
import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  behaviorHint,
  behaviorOf,
  combineDateAndTime,
  displayDateOf,
  formatDateTimeValue,
  isDateOnlyColumn,
  localZoneLabel,
  parseDateTimeValue,
  serializeDateTimeValue,
  withPickedDate,
  withPickedTime,
} from './dateTimeValue';

// Setting TZ at run time switches the zone Date and Intl use.
function inZone(tz: string): void {
  vi.stubEnv('TZ', tz);
}
afterEach(() => {
  vi.unstubAllEnvs();
  vi.restoreAllMocks();
});

const DATE_ONLY = { format: 'DateOnly', dateTimeBehavior: 'DateOnly' } as const;
const TZI = { format: 'DateAndTime', dateTimeBehavior: 'TimeZoneIndependent' } as const;
const TZI_DATE = { format: 'DateOnly', dateTimeBehavior: 'TimeZoneIndependent' } as const;
const USER_LOCAL = { format: 'DateAndTime', dateTimeBehavior: 'UserLocal' } as const;
const USER_LOCAL_DATE = { format: 'DateOnly', dateTimeBehavior: 'UserLocal' } as const;

/** The local fields a picker shows for `d`. */
const local = (d: Date | null) =>
  d && [
    d.getFullYear(),
    d.getMonth() + 1,
    d.getDate(),
    d.getHours(),
    d.getMinutes(),
    d.getSeconds(),
  ];

describe('behaviorOf and isDateOnlyColumn', () => {
  it('defaults to UserLocal when metadata has no behavior', () => {
    expect(behaviorOf({})).toBe('UserLocal');
    expect(behaviorOf({ dateTimeBehavior: 'TimeZoneIndependent' })).toBe('TimeZoneIndependent');
  });

  it('is date-only for the DateOnly format or the DateOnly behavior', () => {
    expect(isDateOnlyColumn({ format: 'DateOnly' })).toBe(true);
    expect(isDateOnlyColumn({ dateTimeBehavior: 'DateOnly' })).toBe(true);
    expect(isDateOnlyColumn(USER_LOCAL)).toBe(false);
    expect(isDateOnlyColumn({})).toBe(false);
  });
});

describe('parseDateTimeValue', () => {
  it('reads empty values as no date', () => {
    expect(parseDateTimeValue(null, USER_LOCAL)).toBeNull();
    expect(parseDateTimeValue(undefined, DATE_ONLY)).toBeNull();
    expect(parseDateTimeValue('', TZI)).toBeNull();
  });

  it('DateOnly: the stored calendar date west of UTC, not the day before (#92)', () => {
    inZone('America/Los_Angeles');
    expect(local(parseDateTimeValue('2026-05-01', DATE_ONLY))).toEqual([2026, 5, 1, 0, 0, 0]);
    // A timestamp on a DateOnly column still means that calendar date.
    expect(local(parseDateTimeValue('2026-05-01T00:00:00Z', DATE_ONLY))).toEqual([
      2026, 5, 1, 0, 0, 0,
    ]);
  });

  it('DateOnly: the same date east of UTC', () => {
    inZone('Pacific/Kiritimati');
    expect(local(parseDateTimeValue('2026-12-31', DATE_ONLY))).toEqual([2026, 12, 31, 0, 0, 0]);
  });

  it('DateOnly: a day whose local midnight is skipped by DST is still that day', () => {
    // Chile springs forward at 00:00 on 2026-09-06, so local midnight is 01:00.
    inZone('America/Santiago');
    const d = parseDateTimeValue('2026-09-06', DATE_ONLY);
    expect(d && [d.getFullYear(), d.getMonth() + 1, d.getDate()]).toEqual([2026, 9, 6]);
    expect(serializeDateTimeValue(d!, DATE_ONLY)).toBe('2026-09-06');
  });

  it('rejects impossible dates instead of rolling them over', () => {
    inZone('UTC');
    expect(parseDateTimeValue('2026-02-31', DATE_ONLY)).toBeNull();
    expect(parseDateTimeValue('2026-02-29', DATE_ONLY)).toBeNull();
    expect(parseDateTimeValue('2026-13-01', TZI)).toBeNull();
    expect(local(parseDateTimeValue('2028-02-29', DATE_ONLY))).toEqual([2028, 2, 29, 0, 0, 0]);
  });

  it('rejects text that is not a date', () => {
    expect(parseDateTimeValue('not a date', DATE_ONLY)).toBeNull();
    expect(parseDateTimeValue('01/05/2026', TZI)).toBeNull();
    expect(parseDateTimeValue('not a date', USER_LOCAL)).toBeNull();
  });

  it('TimeZoneIndependent: the wall-clock digits, ignoring Z and offsets', () => {
    inZone('Asia/Kolkata');
    expect(local(parseDateTimeValue('2026-05-01T23:30:15Z', TZI))).toEqual([
      2026, 5, 1, 23, 30, 15,
    ]);
    expect(local(parseDateTimeValue('2026-05-01T23:30:15+05:00', TZI))).toEqual([
      2026, 5, 1, 23, 30, 15,
    ]);
    inZone('America/Los_Angeles');
    expect(local(parseDateTimeValue('2026-05-01T23:30:15Z', TZI))).toEqual([
      2026, 5, 1, 23, 30, 15,
    ]);
  });

  it('TimeZoneIndependent: optional seconds, a space separator, and a bare date', () => {
    inZone('UTC');
    expect(local(parseDateTimeValue('2026-05-01T08:15', TZI))).toEqual([2026, 5, 1, 8, 15, 0]);
    expect(local(parseDateTimeValue('2026-05-01 08:15:09', TZI))).toEqual([2026, 5, 1, 8, 15, 9]);
    expect(local(parseDateTimeValue('2026-05-01', TZI))).toEqual([2026, 5, 1, 0, 0, 0]);
  });

  it('UserLocal: an instant, shown in the local zone', () => {
    inZone('Asia/Kolkata');
    const d = parseDateTimeValue('2026-05-01T12:00:00Z', USER_LOCAL);
    expect(d?.getTime()).toBe(Date.UTC(2026, 4, 1, 12, 0, 0));
    expect(local(d)).toEqual([2026, 5, 1, 17, 30, 0]);
  });

  it('UserLocal: a bare date is that local calendar date, not UTC midnight (#92)', () => {
    inZone('America/New_York');
    expect(local(parseDateTimeValue('2026-05-01', USER_LOCAL))).toEqual([2026, 5, 1, 0, 0, 0]);
    expect(local(parseDateTimeValue('2026-05-01', {}))).toEqual([2026, 5, 1, 0, 0, 0]);
  });
});

describe('serializeDateTimeValue', () => {
  it('DateOnly: the local date as YYYY-MM-DD, however late in the day', () => {
    inZone('America/Los_Angeles');
    expect(serializeDateTimeValue(new Date(2026, 4, 1, 23, 59), DATE_ONLY)).toBe('2026-05-01');
    inZone('Pacific/Kiritimati');
    expect(serializeDateTimeValue(new Date(2026, 0, 1, 0, 0), DATE_ONLY)).toBe('2026-01-01');
  });

  it('TimeZoneIndependent: the digits picked, as …Z, with no conversion', () => {
    inZone('Asia/Kolkata');
    expect(serializeDateTimeValue(new Date(2026, 4, 1, 23, 30, 5), TZI)).toBe(
      '2026-05-01T23:30:05Z',
    );
    inZone('America/Los_Angeles');
    expect(serializeDateTimeValue(new Date(2026, 4, 1, 7, 5, 0), TZI)).toBe('2026-05-01T07:05:00Z');
  });

  it('TimeZoneIndependent with the DateOnly format: midnight Z, whatever time is set', () => {
    inZone('Asia/Kolkata');
    expect(serializeDateTimeValue(new Date(2026, 4, 1, 15, 45), TZI_DATE)).toBe(
      '2026-05-01T00:00:00Z',
    );
  });

  it('UserLocal: local time converted to UTC, without milliseconds (#106)', () => {
    inZone('Asia/Kolkata');
    expect(serializeDateTimeValue(new Date(2026, 4, 1, 9, 0, 0, 123), USER_LOCAL)).toBe(
      '2026-05-01T03:30:00Z',
    );
    expect(serializeDateTimeValue(new Date(2026, 4, 1, 3, 0), {})).toBe('2026-04-30T21:30:00Z');
  });

  it('UserLocal with the DateOnly format: local midnight as a UTC instant', () => {
    inZone('America/Los_Angeles');
    expect(serializeDateTimeValue(new Date(2026, 4, 1, 15, 0), USER_LOCAL_DATE)).toBe(
      '2026-05-01T07:00:00Z',
    );
    inZone('Asia/Kolkata');
    expect(serializeDateTimeValue(new Date(2026, 4, 1, 15, 0), USER_LOCAL_DATE)).toBe(
      '2026-04-30T18:30:00Z',
    );
  });

  it('UserLocal: uses the offset in force on that date, either side of a DST change', () => {
    inZone('America/New_York');
    // 2026-03-08: clocks go from 02:00 EST to 03:00 EDT.
    expect(serializeDateTimeValue(new Date(2026, 2, 8, 1, 30), USER_LOCAL)).toBe(
      '2026-03-08T06:30:00Z',
    );
    expect(serializeDateTimeValue(new Date(2026, 2, 8, 3, 30), USER_LOCAL)).toBe(
      '2026-03-08T07:30:00Z',
    );
    // 2026-11-01: back from 02:00 EDT to 01:00 EST.
    expect(serializeDateTimeValue(new Date(2026, 10, 1, 0, 30), USER_LOCAL)).toBe(
      '2026-11-01T04:30:00Z',
    );
    expect(serializeDateTimeValue(new Date(2026, 10, 1, 3, 0), USER_LOCAL)).toBe(
      '2026-11-01T08:00:00Z',
    );
  });
});

describe('parse then serialize sends the stored value back unchanged', () => {
  const cases: Array<[string, Parameters<typeof parseDateTimeValue>[1], string]> = [
    ['DateOnly', DATE_ONLY, '2026-05-01'],
    ['TimeZoneIndependent', TZI, '2026-05-01T23:30:00Z'],
    ['TimeZoneIndependent, DateOnly format', TZI_DATE, '2026-05-01T00:00:00Z'],
    ['UserLocal', USER_LOCAL, '2026-05-01T23:30:00Z'],
  ];
  for (const tz of ['America/Los_Angeles', 'UTC', 'Asia/Kolkata', 'Pacific/Kiritimati']) {
    for (const [label, col, value] of cases) {
      it(`${label} in ${tz}`, () => {
        inZone(tz);
        expect(serializeDateTimeValue(parseDateTimeValue(value, col)!, col)).toBe(value);
      });
    }
  }

  it('UserLocal, DateOnly format, read in the zone it was saved from', () => {
    // Local midnight in New York (EDT) is 04:00Z.
    inZone('America/New_York');
    const v = '2026-05-01T04:00:00Z';
    expect(serializeDateTimeValue(parseDateTimeValue(v, USER_LOCAL_DATE)!, USER_LOCAL_DATE)).toBe(
      v,
    );
  });
});

describe('combineDateAndTime', () => {
  it('takes the date from one value and the hours and minutes from the other', () => {
    inZone('Asia/Kolkata');
    const out = combineDateAndTime(
      new Date(2026, 4, 1, 22, 10, 30),
      new Date(1999, 0, 2, 8, 45, 59),
    );
    expect(local(out)).toEqual([2026, 5, 1, 8, 45, 0]);
  });

  it('is midnight with no time', () => {
    inZone('America/Los_Angeles');
    expect(local(combineDateAndTime(new Date(2026, 4, 1, 22, 10), null))).toEqual([
      2026, 5, 1, 0, 0, 0,
    ]);
  });
});

describe('localZoneLabel', () => {
  it('names the zone and its offset on that date', () => {
    inZone('Asia/Kolkata');
    expect(localZoneLabel(new Date(2026, 0, 1))).toMatch(/^Asia\/(Kolkata|Calcutta), UTC\+05:30$/);
  });

  it('follows daylight saving', () => {
    inZone('America/New_York');
    expect(localZoneLabel(new Date(2026, 0, 15))).toBe('America/New_York, UTC-05:00');
    expect(localZoneLabel(new Date(2026, 6, 15))).toBe('America/New_York, UTC-04:00');
  });

  it('falls back to the offset alone when the zone name is unavailable', () => {
    inZone('Asia/Kolkata');
    vi.spyOn(Intl, 'DateTimeFormat').mockImplementation(() => {
      throw new Error('no Intl');
    });
    expect(localZoneLabel(new Date(2026, 0, 1))).toBe('UTC+05:30');
  });
});

describe('behaviorHint', () => {
  it('describes each behavior', () => {
    expect(behaviorHint(DATE_ONLY)).toMatch(/^Date only — sent as YYYY-MM-DD/);
    expect(behaviorHint(TZI_DATE)).toMatch(/sent as that date at 00:00:00Z/);
    expect(behaviorHint(TZI)).toMatch(/sent exactly as entered/);
  });

  it('names the browser zone for UserLocal', () => {
    inZone('America/New_York');
    expect(behaviorHint(USER_LOCAL, new Date(2026, 6, 1))).toContain(
      '(America/New_York, UTC-04:00)',
    );
    expect(behaviorHint({}, new Date(2026, 0, 1))).toContain('(America/New_York, UTC-05:00)');
  });
});

describe('formatDateTimeValue', () => {
  it('shows a DateOnly value as that local date', () => {
    inZone('America/Los_Angeles');
    expect(formatDateTimeValue('2026-05-01', DATE_ONLY)).toBe(
      new Date(2026, 4, 1).toLocaleDateString(),
    );
  });

  it('shows a TimeZoneIndependent value with its own digits', () => {
    inZone('Asia/Kolkata');
    expect(formatDateTimeValue('2026-05-01T23:30:00Z', TZI)).toBe(
      new Date(2026, 4, 1, 23, 30).toLocaleString(),
    );
  });

  it('returns a value it cannot read as it is', () => {
    expect(formatDateTimeValue('garbage', USER_LOCAL)).toBe('garbage');
    expect(formatDateTimeValue('2026-02-31', DATE_ONLY)).toBe('2026-02-31');
  });
});

describe('a time zone independent time in the DST gap (#337)', () => {
  // 02:30 on 8 March 2026 doesn't exist in New York (clocks go 02:00 → 03:00).
  const GAP = '2026-03-08T02:30:00Z';

  it('is shown with its stored digits', () => {
    inZone('America/New_York');
    const shown = displayDateOf(GAP, TZI)!;
    expect(shown.timeZone).toBe('UTC');
    expect(shown.date.toISOString()).toBe('2026-03-08T02:30:00.000Z');
    expect(formatDateTimeValue(GAP, TZI)).toBe(
      new Date(Date.UTC(2026, 2, 8, 2, 30)).toLocaleString(undefined, { timeZone: 'UTC' }),
    );
  });

  it("isn't rewritten when the time picker reports the time it shows (on blur)", () => {
    inZone('America/New_York');
    // The picker shows 03:30, the local Date the gap time became.
    expect(withPickedTime(GAP, new Date(2026, 2, 8, 3, 30), TZI)).toBeNull();
    // A different time is a deliberate pick.
    expect(withPickedTime(GAP, new Date(2026, 2, 8, 4, 15), TZI)).toBe('2026-03-08T04:15:00Z');
  });

  it('keeps its stored time when only the date changes', () => {
    inZone('America/New_York');
    expect(withPickedDate(GAP, new Date(2026, 2, 9), TZI)).toBe('2026-03-09T02:30:00Z');
  });
});

describe('withPickedDate', () => {
  it('keeps the time already set, per behavior', () => {
    inZone('America/New_York');
    expect(withPickedDate('2026-05-01T23:30:00Z', new Date(2026, 4, 2), TZI)).toBe(
      '2026-05-02T23:30:00Z',
    );
    // 15:45Z is 11:45 in New York on both days; the instant moves a day.
    expect(withPickedDate('2026-03-08T15:45:00Z', new Date(2026, 2, 9), USER_LOCAL)).toBe(
      '2026-03-09T15:45:00Z',
    );
  });

  it('sends midnight when there was no value, and a date for date-only columns', () => {
    inZone('America/New_York');
    expect(withPickedDate(undefined, new Date(2026, 2, 9), TZI)).toBe('2026-03-09T00:00:00Z');
    expect(withPickedDate('2026-03-08', new Date(2026, 2, 9), DATE_ONLY)).toBe('2026-03-09');
    expect(withPickedDate('2026-03-08T00:00:00Z', new Date(2026, 2, 9), TZI_DATE)).toBe(
      '2026-03-09T00:00:00Z',
    );
  });
});

describe('withPickedTime', () => {
  it('sets the time on the date already picked, per behavior', () => {
    inZone('America/New_York');
    expect(withPickedTime('2026-05-01T23:30:00Z', new Date(2026, 0, 1, 8, 5), TZI)).toBe(
      '2026-05-01T08:05:00Z',
    );
    // 08:05 in New York on 1 May is 12:05 UTC.
    expect(withPickedTime('2026-05-01T16:00:00Z', new Date(2026, 0, 1, 8, 5), USER_LOCAL)).toBe(
      '2026-05-01T12:05:00Z',
    );
  });

  it('changes nothing without a date, or for the time already shown', () => {
    inZone('America/New_York');
    expect(withPickedTime(undefined, new Date(2026, 0, 1, 8, 5), TZI)).toBeNull();
    expect(withPickedTime('2026-05-01T08:05:42Z', new Date(2026, 0, 1, 8, 5), TZI)).toBeNull();
  });
});
