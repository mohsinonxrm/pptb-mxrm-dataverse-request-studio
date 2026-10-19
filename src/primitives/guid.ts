/**
 * Record ids (GUIDs) as Dataverse expects them in an OData key segment:
 * `accounts(00000000-0000-0000-0000-000000000000)`. The Web API only parses
 * the dashed 8-4-4-4-12 form there, so ids that reach a URL must be in it.
 */

const DASHED_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const COMPACT_RE = /^[0-9a-f]{32}$/i;

/**
 * Parse a GUID typed or pasted by the user and return it in the lowercase
 * 8-4-4-4-12 form, or null when it isn't a GUID.
 *
 * Accepted: the dashed form and the 32-digit dash-less form (as SQL and
 * some exports print it), either one optionally wrapped in `{}` or `()`,
 * with surrounding whitespace. Partly dashed values are rejected — they
 * are more likely a typo than a deliberate format.
 */
export function parseGuid(input: string | null | undefined): string | null {
  if (!input) return null;
  let s = input.trim();
  const m = /^\{(.*)\}$/.exec(s) ?? /^\((.*)\)$/.exec(s);
  if (m) s = m[1].trim();
  if (DASHED_RE.test(s)) return s.toLowerCase();
  if (COMPACT_RE.test(s)) {
    const h = s.toLowerCase();
    return `${h.slice(0, 8)}-${h.slice(8, 12)}-${h.slice(12, 16)}-${h.slice(16, 20)}-${h.slice(20)}`;
  }
  return null;
}

/**
 * True when `s` is a record id that can go into an OData key as-is (the
 * dashed form, no braces or whitespace). This is what the modes check
 * before they build a URL.
 */
export function isRecordGuid(s: string | null | undefined): boolean {
  return !!s && DASHED_RE.test(s);
}
