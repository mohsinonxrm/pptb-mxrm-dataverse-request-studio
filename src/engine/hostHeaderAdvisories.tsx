// Host header-limitation advisories — blocks Execute when the request relies
// on a header the PPTB host won't transmit.
//
// PPTB 1.2.7+ sends the write modes' composed headers as `additionalHeaders`
// (see dataverseExecutor.ts), so nothing here fires on those hosts. Older
// hosts' `dataverseAPI` has no header parameter: precondition (`If-Match` /
// `If-None-Match`) and bypass (`MSCRM.*`) headers would be dropped at execute
// time — an Upsert set to "create-only" would overwrite an existing row, an
// ETag check would never fail, and plug-ins would still run (#45). There we
// refuse to execute those requests and point the user at the Code tab.
//
// The Update default (`If-Match: *`, update-only) stays executable on older
// hosts: it's the default for every Update.

import type { BypassOptions, ConcurrencyMode } from '../state/writeState';
import { adv, type Advisory } from '../primitives/advisories';
import { hostSupportsAdditionalHeaders, isEmbedded } from '../host/pptbBridge';

export type HeaderSensitiveOperation = 'create' | 'update' | 'upsert' | 'delete' | 'merge';

const COPY_HINT =
  'Update Power Platform ToolBox, or copy the request from the Code tab to run it outside it.';

/** True when an older PPTB (no header pass-through) would silently drop a
 *  precondition the request depends on — the same rule gates Execute here
 *  and the Precondition pane's notice. */
export function isConcurrencyEnforcementRequired(
  op: HeaderSensitiveOperation,
  concurrency: ConcurrencyMode,
): boolean {
  switch (concurrency.kind) {
    case 'none':
      return false;
    case 'etag':
    case 'create-only':
      return true;
    case 'update-only':
      // Upsert's update-only exists to stop the PATCH from creating a row.
      return op === 'upsert';
  }
}

function isBypassActive(b: BypassOptions): boolean {
  return b.businessLogic !== 'none' || b.suppressFlows;
}

export function detectUnsentHeaderAdvisories(opts: {
  op: HeaderSensitiveOperation;
  concurrency?: ConcurrencyMode;
  bypass?: BypassOptions;
  /** Legacy Delete flag — `MSCRM.BypassCustomPluginExecution`. */
  bypassCustomPlugins?: boolean;
  preconditionFocus?: string;
  bypassFocus?: string;
}): Advisory[] {
  if (!isEmbedded() || hostSupportsAdditionalHeaders()) return [];
  const out: Advisory[] = [];

  if (opts.concurrency && isConcurrencyEnforcementRequired(opts.op, opts.concurrency)) {
    out.push(
      adv.err(
        'host-precondition-unsent',
        'header',
        "PPTB can't send this precondition header",
        <span>
          This version of Power Platform ToolBox can't send request headers (1.2.7 and later can),
          so the <code>If-Match</code> / <code>If-None-Match</code> condition would be dropped and
          the request would run without it. Set the precondition to <strong>None</strong> to execute
          here. {COPY_HINT}
        </span>,
        opts.preconditionFocus ?? 'precondition',
      ),
    );
  }

  if ((opts.bypass && isBypassActive(opts.bypass)) || opts.bypassCustomPlugins) {
    out.push(
      adv.err(
        'host-bypass-unsent',
        'header',
        "PPTB can't send the bypass headers",
        <span>
          This version of Power Platform ToolBox can't send request headers (1.2.7 and later can),
          so the <code>MSCRM.*</code> bypass headers would be dropped and plug-ins / flows would
          still run. Turn the bypass options off to execute here. {COPY_HINT}
        </span>,
        opts.bypassFocus ?? 'bypass',
      ),
    );
  }

  return out;
}
