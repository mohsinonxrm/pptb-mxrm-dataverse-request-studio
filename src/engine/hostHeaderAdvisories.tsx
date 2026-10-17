// Host header-limitation advisories — blocks Execute when the request relies
// on a header the PPTB host won't transmit.
//
// The live executor routes every write through `dvHost.create/update/delete/
// execute`, which take no headers (see dataverseExecutor.ts). Precondition
// (`If-Match` / `If-None-Match`) and bypass (`MSCRM.*`) headers are composed
// for the URL bar and Code tab but silently dropped at execute time — so an
// Upsert set to "create-only" would overwrite an existing row, an ETag check
// would never fail, and plug-ins would still run (#45). Until the host
// accepts additional headers (PPTB 1.2.7, #43), we refuse to execute those
// requests in PPTB and point the user at the Code tab instead.
//
// The Update default (`If-Match: *`, update-only) stays executable: it's the
// default for every Update and is tracked with the header pass-through work.

import type { BypassOptions, ConcurrencyMode } from '../state/writeState';
import { adv, type Advisory } from '../primitives/advisories';
import { isEmbedded } from '../host/pptbBridge';

export type HeaderSensitiveOperation = 'create' | 'update' | 'upsert' | 'delete' | 'merge';

const COPY_HINT = 'Copy the request from the Code tab to run it outside Power Platform ToolBox.';

/** True when PPTB would silently drop a precondition the request depends on
 *  — the same rule gates Execute here and the Precondition pane's notice. */
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
  if (!isEmbedded()) return [];
  const out: Advisory[] = [];

  if (opts.concurrency && isConcurrencyEnforcementRequired(opts.op, opts.concurrency)) {
    out.push(
      adv.err(
        'host-precondition-unsent',
        'header',
        "PPTB can't send this precondition header",
        <span>
          Power Platform ToolBox doesn't pass request headers through yet, so the{' '}
          <code>If-Match</code> / <code>If-None-Match</code> condition would be dropped and the
          request would run without it. Set the precondition to <strong>None</strong> to execute
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
          Power Platform ToolBox doesn't pass request headers through yet, so the{' '}
          <code>MSCRM.*</code> bypass headers would be dropped and plug-ins / flows would still run.
          Turn the bypass options off to execute here. {COPY_HINT}
        </span>,
        opts.bypassFocus ?? 'bypass',
      ),
    );
  }

  return out;
}
