// DateTime column values — parse for the pickers, serialize for the wire.
//
// Dataverse has three DateTime behaviors, and each needs its own handling
// (https://learn.microsoft.com/power-apps/maker/data-platform/behavior-format-date-time-field):
//
//   DateOnly             — Web API type Edm.Date. Sent as `YYYY-MM-DD`. No
//                          time, no time zone. Shown as that calendar date.
//   TimeZoneIndependent  — Edm.DateTimeOffset, but stored as typed (no
//                          conversion). Sent as `YYYY-MM-DDTHH:mm:ssZ` with
//                          the wall-clock digits the user picked.
//   UserLocal (default)  — Edm.DateTimeOffset, an instant. The pickers show
//                          the browser's local time; the value is converted
//                          to UTC.
//
// The Format (DateOnly vs DateAndTime) only decides whether a time picker
// shows. A DateOnly-format column with UserLocal / TimeZoneIndependent
// behavior is still Edm.DateTimeOffset, so it gets a full timestamp at
// midnight (local midnight → UTC for UserLocal, `T00:00:00Z` for TZI), like
// model-driven apps do.
//
// The pickers work with JS Dates whose LOCAL fields (getFullYear, getHours…)
// are what the user sees. `new Date('2026-05-01')` is UTC midnight, which
// west of UTC is the previous local day — so never parse stored values with
// the Date constructor except for UserLocal instants (#92).
//
// A local Date can't hold a time in the browser's DST gap: a time zone
// independent 02:30 on a spring-forward day becomes 03:30 (#337). Display
// therefore reads those digits as UTC (`displayDateOf`), and a date change
// keeps the stored time (`withPickedDate`). Fluent's time picker only takes
// a local Date, so it still shows such a time an hour on.

import type { DateTimeBehavior, DateTimeFormat } from '../mock/metadata';

interface DateTimeColumnLike {
  format?: DateTimeFormat;
  dateTimeBehavior?: DateTimeBehavior;
}

const pad = (n: number, len = 2) => String(n).padStart(len, '0');

/** Behavior with Dataverse's default when metadata didn't say. */
export const behaviorOf = (col: DateTimeColumnLike): DateTimeBehavior =>
  col.dateTimeBehavior ?? 'UserLocal';

/** True when the column takes a date with no time picker. */
export const isDateOnlyColumn = (col: DateTimeColumnLike): boolean =>
  col.format === 'DateOnly' || col.dateTimeBehavior === 'DateOnly';

const PARTS_RE = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?/;

function localFromParts(value: string, withTime: boolean): Date | null {
  const m = PARTS_RE.exec(value.trim());
  if (!m) return null;
  const [, y, mo, d, h, mi, s] = m;
  const out = withTime
    ? new Date(+y, +mo - 1, +d, +(h ?? 0), +(mi ?? 0), +(s ?? 0))
    : new Date(+y, +mo - 1, +d);
  // Reject rollovers like 2026-02-31.
  return out.getMonth() === +mo - 1 && out.getDate() === +d ? out : null;
}

/**
 * Parse a stored value into a Date whose local fields are what the pickers
 * should show. Returns null for empty / unparseable values.
 */
export function parseDateTimeValue(
  value: string | null | undefined,
  col: DateTimeColumnLike,
): Date | null {
  if (!value) return null;
  const behavior = behaviorOf(col);
  // DateOnly: the calendar date, wherever the browser is.
  if (behavior === 'DateOnly') return localFromParts(value, false);
  // TimeZoneIndependent: the wall-clock digits, ignoring any `Z`/offset.
  if (behavior === 'TimeZoneIndependent') return localFromParts(value, true);
  // UserLocal: an instant. A bare `YYYY-MM-DD` has no instant — read it as
  // that local calendar date rather than UTC midnight.
  if (/^\d{4}-\d{2}-\d{2}$/.test(value.trim())) return localFromParts(value, false);
  const d = new Date(value);
  return Number.isNaN(d.getTime()) ? null : d;
}

/**
 * Serialize a picker Date (its local fields are what the user picked) into
 * the value sent to Dataverse for this column.
 */
export function serializeDateTimeValue(d: Date, col: DateTimeColumnLike): string {
  const behavior = behaviorOf(col);
  const dateOnly = isDateOnlyColumn(col);
  const ymd = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  if (behavior === 'DateOnly') return ymd;
  if (behavior === 'TimeZoneIndependent') {
    if (dateOnly) return `${ymd}T00:00:00Z`;
    return `${ymd}T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}Z`;
  }
  const instant = dateOnly ? new Date(d.getFullYear(), d.getMonth(), d.getDate()) : d;
  return instant.toISOString().replace(/\.\d{3}Z$/, 'Z');
}

/** Local date of `date` with the local time of `time` (or midnight). */
export function combineDateAndTime(date: Date, time: Date | null): Date {
  return new Date(
    date.getFullYear(),
    date.getMonth(),
    date.getDate(),
    time ? time.getHours() : 0,
    time ? time.getMinutes() : 0,
    0,
  );
}

/** The browser's time zone and its UTC offset on `at`, e.g. "Asia/Kolkata, UTC+05:30". */
export function localZoneLabel(at: Date = new Date()): string {
  let name = '';
  try {
    name = Intl.DateTimeFormat().resolvedOptions().timeZone ?? '';
  } catch {
    /* older engines — offset alone is enough */
  }
  const off = -at.getTimezoneOffset();
  const sign = off >= 0 ? '+' : '-';
  const abs = Math.abs(off);
  const utc = `UTC${sign}${pad(Math.floor(abs / 60))}:${pad(abs % 60)}`;
  return name ? `${name}, ${utc}` : utc;
}

/** One-line description of what the picker shows and what gets sent. */
export function behaviorHint(col: DateTimeColumnLike, at?: Date | null): string {
  const behavior = behaviorOf(col);
  if (behavior === 'DateOnly') return 'Date only — sent as YYYY-MM-DD, no time zone conversion.';
  if (behavior === 'TimeZoneIndependent') {
    return isDateOnlyColumn(col)
      ? 'Time zone independent — sent as that date at 00:00:00Z, no conversion.'
      : 'Time zone independent — the date and time are sent exactly as entered (as …Z), no conversion.';
  }
  return (
    `User local — you enter local time in your browser's zone (${localZoneLabel(at ?? undefined)}); ` +
    'it is converted to UTC. Dataverse shows it in each user’s own time zone setting.'
  );
}

/**
 * What to format a stored value with: a Date, and for a time zone
 * independent value `timeZone: 'UTC'`, so its digits show as stored even
 * when they fall in the browser's DST gap (#337). Null when unreadable.
 */
export function displayDateOf(
  value: string,
  col: DateTimeColumnLike,
): { date: Date; timeZone?: 'UTC' } | null {
  if (behaviorOf(col) === 'TimeZoneIndependent') {
    const m = PARTS_RE.exec(value.trim());
    if (!m) return null;
    const [, y, mo, d, h, mi, s] = m;
    const date = new Date(Date.UTC(+y, +mo - 1, +d, +(h ?? 0), +(mi ?? 0), +(s ?? 0)));
    // Reject rollovers like 2026-02-31.
    if (date.getUTCMonth() !== +mo - 1 || date.getUTCDate() !== +d) return null;
    return { date, timeZone: 'UTC' };
  }
  const date = parseDateTimeValue(value, col);
  return date ? { date } : null;
}

/** Display text for a stored value (diff / merge panes), without shifting it. */
export function formatDateTimeValue(value: string, col: DateTimeColumnLike): string {
  const shown = displayDateOf(value, col);
  if (!shown) return value;
  const opts = { timeZone: shown.timeZone };
  return isDateOnlyColumn(col)
    ? shown.date.toLocaleDateString(undefined, opts)
    : shown.date.toLocaleString(undefined, opts);
}

/**
 * The value after the date picker picks `date`, keeping the time of day
 * already set (seconds drop to 00, as the pickers don't show them). A time
 * zone independent value keeps its stored hour and minute, which its local
 * Date may have moved out of a DST gap (#337).
 */
export function withPickedDate(
  value: string | undefined,
  date: Date,
  col: DateTimeColumnLike,
): string {
  const dateOnly = isDateOnlyColumn(col);
  if (behaviorOf(col) === 'TimeZoneIndependent' && !dateOnly) {
    const m = value ? PARTS_RE.exec(value.trim()) : null;
    const time = m?.[4] ? `${m[4]}:${m[5]}:00` : '00:00:00';
    return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}T${time}Z`;
  }
  const current = parseDateTimeValue(value, col);
  return serializeDateTimeValue(combineDateAndTime(date, dateOnly ? null : current), col);
}

/**
 * The value after the time picker picks `time`, or null when nothing should
 * change: no date yet, or the time already shown. The picker reports its
 * time on blur too, and a time in a DST gap is shown an hour on, so sending
 * the shown time back would move it (#337).
 */
export function withPickedTime(
  value: string | undefined,
  time: Date,
  col: DateTimeColumnLike,
): string | null {
  const current = parseDateTimeValue(value, col);
  if (!current) return null;
  if (time.getHours() === current.getHours() && time.getMinutes() === current.getMinutes()) {
    return null;
  }
  return serializeDateTimeValue(combineDateAndTime(current, time), col);
}
