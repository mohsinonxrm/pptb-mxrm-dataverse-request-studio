// The headers a mode sends, composed in one place (#244).
//
// A mode's own header rows, plus the rows its panes compose: the precondition
// (If-Match / If-None-Match), duplicate detection, and the bypass family.
// The write modes (Create, Update, Upsert, Delete, Merge, Bulk Write) pass
// their whole state here for Execute and the Code tab, Retrieve Multiple for
// its Prefer, and the collection exports and $batch for every saved request,
// so a fix to one rule (#217, #218, #264) reaches all of them.

import { applyBypassToHeaders } from './bypassHeaders';
import { COUNT_ONLY_PREFER } from './urlBuilder';
import { headerItemsToObject, type HeaderItem } from '../editors/HeadersEditor';
import { preferToHeaderString, type PreferSpec } from '../editors/PreferEditor';
import { preconditionToHeader } from '../editors/PreconditionEditor';
import type { BypassOptions, ConcurrencyMode } from '../state/writeState';

/** The header-related fields a mode state may carry. */
export interface HeaderBits {
  headers?: HeaderItem[];
  prefer?: PreferSpec;
  /** Retrieve Multiple count only (#272): its own Prefer, never a page size. */
  countOnly?: boolean;
  concurrency?: ConcurrencyMode;
  bypass?: BypassOptions;
  /** Delete's legacy "bypass custom plugins" switch. */
  bypassCustomPlugins?: boolean;
  /** Create: duplicate detection on → MSCRM.SuppressDuplicateDetection: false. */
  duplicateDetection?: boolean;
  /** Merge: skip duplicate detection on UpdateContent. */
  suppressDuplicateDetection?: boolean;
}

function composed(name: string, value: string, hint: string): HeaderItem {
  return { id: `__composed__${name}`, name, value, enabled: true, builtin: true, hint };
}

/**
 * The header rows a mode sends, in order: its own rows (Create's duplicate
 * detection row follows its switch), the precondition, Merge's
 * duplicate-detection skip, then the bypass headers.
 */
export function composeHeaderItems(s: HeaderBits): HeaderItem[] {
  let h: HeaderItem[] = (s.headers ?? []).map((it) =>
    it.name === 'MSCRM.SuppressDuplicateDetection' && s.duplicateDetection !== undefined
      ? { ...it, enabled: s.duplicateDetection, value: 'false' }
      : it,
  );
  const cc = s.concurrency ? preconditionToHeader(s.concurrency) : null;
  if (cc) h.push(composed(cc.name, cc.value, 'Composed from the Optimistic concurrency pane.'));
  if (s.suppressDuplicateDetection) {
    h.push(
      composed(
        'MSCRM.SuppressDuplicateDetection',
        'true',
        'Skips duplicate detection rules when applying UpdateContent overrides.',
      ),
    );
  }
  if (s.bypass) {
    // Delete's legacy switch sends the legacy plug-in bypass when the Bypass
    // pane itself is off.
    const b =
      s.bypassCustomPlugins && s.bypass.businessLogic === 'none'
        ? { ...s.bypass, businessLogic: 'sync' as const, useLegacyHeader: true }
        : s.bypass;
    h = applyBypassToHeaders(h, b);
  }
  return h;
}

/** The request's headers by name, with the Prefer its pane composes. */
export function composeHeaders(s: HeaderBits): Record<string, string> {
  return headerItemsToObject(
    composeHeaderItems(s),
    s.countOnly ? COUNT_ONLY_PREFER : s.prefer ? preferToHeaderString(s.prefer) : null,
  );
}
