/**
 * Checks for the values DRS puts into a URL path as-is: record ids in key
 * segments (`accounts(<id>)`) and names in path segments (`/<navProperty>`).
 * Values come from the pickers, but also from typed GUIDs and from imported
 * saved requests, which are untrusted files. Without a check, a value like
 * `x)/other_nav?$filter=…` would send a different request from the one the
 * user meant (#221).
 */

import { isRecordGuid } from './guid';

const IDENTIFIER_RE = /^[A-Za-z_][A-Za-z0-9_]*$/;
const PROPERTY_PATH_RE = /^[A-Za-z_][A-Za-z0-9_]*(\/[A-Za-z_][A-Za-z0-9_]*)*$/;

/** An ISO 8601 date or date-time, as OData writes it unquoted. */
export const ISO_DATE_RE =
  /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2}(\.\d+)?)?(Z|[+-]\d{2}:\d{2})?)?$/;

/** True when `s` is an OData simple identifier (a navigation property, column or table name). */
export const isODataIdentifier = (s: string | null | undefined): boolean =>
  !!s && IDENTIFIER_RE.test(s);

/**
 * True when `s` is a property path: identifiers joined by `/` (a column, a
 * lambda alias + column, a lookup's nav + related column).
 */
export const isPropertyPath = (s: string | null | undefined): boolean =>
  !!s && PROPERTY_PATH_RE.test(s);

const shown = (v: string) => JSON.stringify(v.length > 40 ? `${v.slice(0, 40)}…` : v);

type Part = readonly [label: string, value: string | null | undefined];

/**
 * Why one of these record ids can't go into a URL, or null when they all can.
 * Empty values pass: each mode asks for a missing id with its own wording.
 */
export function invalidIdReason(ids: readonly Part[]): string | null {
  for (const [label, value] of ids) {
    if (value && !isRecordGuid(value)) {
      return `${label} ${shown(value)} isn't a GUID (00000000-0000-0000-0000-000000000000).`;
    }
  }
  return null;
}

/** Why one of these names can't go into a URL path, or null when they all can. Empty values pass. */
export function invalidNameReason(names: readonly Part[]): string | null {
  for (const [label, value] of names) {
    if (value && !isODataIdentifier(value)) {
      return `${label} ${shown(value)} isn't a valid name (letters, digits and underscores).`;
    }
  }
  return null;
}

/**
 * A URL with its spaces percent-encoded. The URL bar keeps `$filter`'s spaces
 * literal, to match the Microsoft Learn examples, but a space isn't a valid
 * URI character in an exported collection, generated code or an HTTP action.
 */
export const encodeSpaces = (url: string): string => url.replace(/ /g, '%20');

/** Both checks: the first id problem, else the first name problem. */
export const invalidUrlPartReason = (ids: readonly Part[], names: readonly Part[] = []) =>
  invalidIdReason(ids) ?? invalidNameReason(names);
