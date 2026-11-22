// The URL of a `$batch` part (#225). ToolBox's `executeBatch` takes a path
// relative to `/api/data/v9.2/` and rejects anything else with a generic
// message, so `toBatchUrl` turns what DRS has (a Code tab path, an absolute
// `@odata.nextLink`) into that shape, and refuses what the host would refuse
// with a message that says why.

/**
 * Thrown by `toBatchUrl` for a path that can't go in a `$batch` part. Nothing
 * was sent; `statusCode` makes the executor's `mapError` report it as a 400.
 */
export class BatchPathError extends Error {
  readonly statusCode = 400;
  constructor(reason: string) {
    super(`Not sent: ${reason}`);
    this.name = 'BatchPathError';
  }
}

// `https://<org>/api/data/v9.x/` (absolute) or `/api/data/v9.x/` (relative).
const ABSOLUTE_API_PREFIX_RE = /^https?:\/\/[^/?#]+\/api\/data\/v\d+(?:\.\d+)?\//i;
const RELATIVE_API_PREFIX_RE = /^\/?api\/data\/v\d+(?:\.\d+)?\//i;
// What's left must be a relative path: no scheme, no `//host`, no backslash.
const NOT_RELATIVE_RE = /^(?:[a-z][a-z0-9+.-]*:|\/)|\\/i;

/** Up to three rounds of percent-decoding, as the host checks (a bad escape stops it). */
function decodedForms(segment: string): string[] {
  const forms = [segment];
  let s = segment;
  for (let i = 0; i < 3; i++) {
    try {
      const next = decodeURIComponent(s);
      if (next === s) break;
      forms.push(next);
      s = next;
    } catch {
      break;
    }
  }
  return forms;
}

/**
 * `url` without its `https://<org>/api/data/v9.x/` (or `/api/data/v9.x/`)
 * prefix, so what's left is relative to the Web API base; unchanged when it
 * has neither. Shared by `toBatchUrl` and the typed calls that follow an
 * absolute `@odata.nextLink` / `@odata.deltaLink` (#233).
 */
export function stripApiPrefix(url: string): string {
  if (ABSOLUTE_API_PREFIX_RE.test(url)) return url.replace(ABSOLUTE_API_PREFIX_RE, '');
  return url.replace(RELATIVE_API_PREFIX_RE, '');
}

/**
 * The URL to put on a `$batch` part for `path`:
 *   - a leading `/`, and an `https://<org>/api/data/v9.x/` (or `api/data/v9.x/`)
 *     prefix, are removed; the host adds `api/data/v9.2/` back;
 *   - every literal `?` after the first is written `%3F`: the host splits the
 *     URL with `split("?", 2)`, which drops everything after a second `?`.
 * Throws `BatchPathError` for an empty path, an absolute URL that isn't a
 * Dataverse Web API URL, a `..` segment or a nested `$batch`, all of which
 * the host would reject.
 */
export function toBatchUrl(path: string): string {
  const trimmed = path.trim();
  let p = stripApiPrefix(trimmed);
  if (p === trimmed && p.startsWith('/')) p = p.slice(1);
  if (!p) throw new BatchPathError('the request path is empty.');
  if (NOT_RELATIVE_RE.test(p)) {
    throw new BatchPathError(
      `"${path.trim()}" isn't a path under /api/data/v9.2/, so it can't be sent through $batch.`,
    );
  }

  const q = p.indexOf('?');
  const resource = q >= 0 ? p.slice(0, q) : p;
  const query = q >= 0 ? p.slice(q + 1) : undefined;
  for (const segment of resource.split('/')) {
    const forms = decodedForms(segment);
    if (forms.some((f) => f === '..' || f === '.')) {
      throw new BatchPathError(`the request path can't contain "${segment}" segments.`);
    }
    if (forms.some((f) => f.toLowerCase() === '$batch')) {
      throw new BatchPathError("a $batch request can't be sent inside another $batch.");
    }
  }
  if (!resource) throw new BatchPathError('the request path is empty.');
  return query === undefined ? resource : `${resource}?${query.replace(/\?/g, '%3F')}`;
}
