// Request headers that carry credentials or identify a person (#217).
//
// Credentials (`Authorization`, `Proxy-Authorization`, `Cookie`): DRS never
// sends, saves or exports them. Power Platform ToolBox authenticates every
// request itself and drops these, the Code tab inserts its own token
// placeholder, and an exported collection authenticates through its
// `accessToken` variable. A value typed into the Headers pane would only
// leak into localStorage and into files people share.
//
// Impersonation (`MSCRMCallerID`, `CallerObjectId`): legitimate, so saved
// requests keep them; collection exports write a variable instead of the
// user's id.

/** Credential header names, lower-case. */
export const CREDENTIAL_HEADERS: ReadonlySet<string> = new Set([
  'authorization',
  'proxy-authorization',
  'cookie',
]);

/** True for a header DRS never sends, saves or exports. */
export function isCredentialHeader(name: string): boolean {
  return CREDENTIAL_HEADERS.has(name.trim().toLowerCase());
}

/** Impersonation header (lower-case) → the collection variable that replaces its value. */
export const IMPERSONATION_HEADER_VARS: Readonly<Record<string, string>> = {
  mscrmcallerid: 'callerId',
  callerobjectid: 'callerObjectId',
};

/** The collection variable for an impersonation header, or undefined. */
export function impersonationVar(name: string): string | undefined {
  return IMPERSONATION_HEADER_VARS[name.trim().toLowerCase()];
}

/** A header map without credential headers. */
export function withoutCredentialHeaders(headers: Record<string, string>): Record<string, string> {
  return Object.fromEntries(Object.entries(headers).filter(([name]) => !isCredentialHeader(name)));
}

interface HeaderRowLike {
  name: string;
  value: string;
  enabled: boolean;
}

function isHeaderRow(x: unknown): x is HeaderRowLike {
  return (
    x !== null &&
    typeof x === 'object' &&
    typeof (x as HeaderRowLike).name === 'string' &&
    typeof (x as HeaderRowLike).value === 'string' &&
    typeof (x as HeaderRowLike).enabled === 'boolean'
  );
}

/**
 * A saved mode state with credential header rows removed. Walks the whole
 * snapshot, so it covers every mode's `headers`, Bulk Write / Bulk Delete
 * and each `$batch` part: any array of `{ name, value, enabled }` rows loses
 * the rows named like a credential header. Returns `state` itself when
 * nothing was removed.
 */
export function stripCredentialRows<T>(state: T): T {
  return mapHeaderRows(state, (row) => (isCredentialHeader(row.name) ? null : row));
}

/** The impersonation presets' blank value (HeadersEditor). */
const NO_CALLER = '00000000-0000-0000-0000-000000000000';

/**
 * A mode state for the run history (#226). Credential rows are removed, as
 * for a saved request. Impersonation rows are kept but reset to the disabled
 * blank preset: a history the user never chose to keep doesn't hold who they
 * impersonated, and a restored request has its presets back, off.
 * `cleared` says whether a caller id was actually removed.
 */
export function stripHistoryRows<T>(state: T): { state: T; cleared: boolean } {
  let cleared = false;
  const out = mapHeaderRows(state, (row) => {
    if (isCredentialHeader(row.name)) return null;
    if (!impersonationVar(row.name)) return row;
    const value = row.value.trim();
    if (value && value !== NO_CALLER) cleared = true;
    return row.value === NO_CALLER && !row.enabled
      ? row
      : { ...row, value: NO_CALLER, enabled: false };
  });
  return { state: out, cleared };
}

/** `state` with every header row passed through `fn` (null removes it), at any depth. */
function mapHeaderRows<T>(state: T, fn: (row: HeaderRowLike) => HeaderRowLike | null): T {
  const walk = (v: unknown): unknown => {
    if (Array.isArray(v)) {
      if (v.length > 0 && v.every(isHeaderRow)) {
        let changed = false;
        const kept: HeaderRowLike[] = [];
        for (const row of v) {
          const next = fn(row);
          if (next !== row) changed = true;
          if (next) kept.push(next);
        }
        return changed ? kept : v;
      }
      let changed = false;
      const out = v.map((item) => {
        const w = walk(item);
        if (w !== item) changed = true;
        return w;
      });
      return changed ? out : v;
    }
    if (v !== null && typeof v === 'object') {
      let changed = false;
      const out: Record<string, unknown> = {};
      for (const [k, item] of Object.entries(v)) {
        const w = walk(item);
        if (w !== item) changed = true;
        out[k] = w;
      }
      return changed ? out : v;
    }
    return v;
  };
  return walk(state) as T;
}
