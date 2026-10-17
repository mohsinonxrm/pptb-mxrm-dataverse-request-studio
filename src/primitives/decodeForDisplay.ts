/**
 * `decodeURIComponent` for values rendered in the UI. Values from
 * URLSearchParams are already decoded once; a second decode makes
 * double-encoded values (Dataverse's `$skiptoken` cookie) readable, but a
 * value holding a literal `%` (e.g. `contains(name,'50%')`) isn't valid
 * percent-encoding and throws a URIError — during render that blanks the
 * whole app (#48). Fall back to the value as given.
 */
export function decodeForDisplay(value: string): string {
  try {
    return decodeURIComponent(value);
  } catch {
    return value;
  }
}
