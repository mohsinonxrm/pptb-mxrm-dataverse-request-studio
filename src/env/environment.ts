// Studio environment — the org we're targeting.
//
// PPTB-only: when embedded + connected we resolve to the host's active
// connection via `getSession().environmentUrl`. Outside PPTB we return a
// blank env (host='', apiBase=API_BASE) so URL construction can
// still proceed as a skeleton without crashing on a null host.
//
// Existing call sites that read `ENV.host` / `ENV.apiBase` keep working
// via the Proxy below; they just become reactive to host changes at the
// next render.

import { getSession } from '../host/pptbBridge';

/**
 * The Web API version in every URL DRS writes. It matches the PPTB host's own
 * `DATAVERSE_API_VERSION`: Execute goes through the host, which always calls
 * v9.2, so another version here would only change the Code tab and the URL
 * bar, not what's sent (#224).
 */
export const API_VERSION = 'v9.2';
export const API_BASE = `/api/data/${API_VERSION}`;

export interface StudioEnv {
  /** Short label — env URL host or PPTB environment enum ("Production", "Dev", …). */
  name: string;
  /** Host portion of the URL — e.g. `contoso.crm.dynamics.com`. */
  host: string;
  /** API base path including version — e.g. `/api/data/v9.2`. */
  apiBase: string;
  /** User principal name when available, falls back to a placeholder. */
  user: string;
}

/**
 * Empty fallback used outside PPTB. Host stays blank so any URL we render
 * is clearly an in-progress skeleton (e.g. `https:///api/data/v9.2/...`)
 * rather than a misleading fake org URL.
 */
const EMPTY_ENV: StudioEnv = {
  name: '(no connection)',
  host: '',
  apiBase: API_BASE,
  user: '',
};

/**
 * Resolves to the live PPTB-pushed env when embedded + connected, otherwise
 * EMPTY_ENV. Pure function — call it inside components for proper reactivity.
 */
export function getEnv(): StudioEnv {
  const s = getSession();
  if (!s.embedded || !s.connected || !s.environmentUrl) return EMPTY_ENV;
  const url = s.environmentUrl.replace(/\/$/, '');
  const host = url.replace(/^https?:\/\//, '');
  return {
    // Prefer the explicit environment label ("Production" / "Dev") when present.
    name: s.environment || host,
    host,
    apiBase: API_BASE,
    user: '',
  };
}

/**
 * Backward-compat Proxy so existing imports of `ENV.host` / `ENV.name` still
 * work without refactoring every call site. Each property access resolves
 * via getEnv() so it picks up host-pushed env changes on the next render.
 */
export const ENV: StudioEnv = new Proxy({} as StudioEnv, {
  get(_t, key: string | symbol) {
    const e = getEnv();
    return (e as unknown as Record<string | symbol, unknown>)[key];
  },
  has(_t, key) {
    return key in getEnv();
  },
  ownKeys() {
    return Object.keys(getEnv());
  },
  getOwnPropertyDescriptor(_t, key) {
    return {
      enumerable: true,
      configurable: true,
      value: (getEnv() as unknown as Record<string | symbol, unknown>)[key as string | symbol],
    };
  },
});
