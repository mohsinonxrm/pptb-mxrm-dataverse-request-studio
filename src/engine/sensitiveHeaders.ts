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
  return stripHeaderRows(state, isCredentialHeader);
}

/**
 * A mode state for the run history (#226): without credential headers, and
 * without impersonation headers either. A saved request keeps the user's
 * MSCRMCallerID on purpose; a history the user never chose to keep doesn't.
 */
export function stripHistoryRows<T>(state: T): T {
  return stripHeaderRows(state, (name) => isCredentialHeader(name) || !!impersonationVar(name));
}

/** `state` with every header row whose name `drop` matches removed, at any depth. */
function stripHeaderRows<T>(state: T, drop: (name: string) => boolean): T {
  const walk = (v: unknown): unknown => {
    if (Array.isArray(v)) {
      if (v.length > 0 && v.every(isHeaderRow)) {
        const kept = v.filter((row) => !drop(row.name));
        return kept.length === v.length ? v : kept;
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
