// Advisory model — the shared shape every mode + editor uses to report
// guardrail conditions to the user.
//
// Modes + editors emit an Advisory[] for any of the following:
//   • a Dataverse anti-pattern was detected (slow query risk)
//   • a wildcard rewrite happened silently (transparency)
//   • a header bypass was applied (privilege requirement)
//   • a documented spec violation that *blocks* execute (missing required field)
//
// Rather than flooding the screen with MessageBars, every mode aggregates
// these into an Advisory[] and renders them in a single collapsible drawer
// near the Execute button. Blockers also surface inline at the source of
// the problem so the user knows where to go fix it.
//
// References:
//   - https://learn.microsoft.com/en-us/power-apps/developer/data-platform/wildcard-characters
//   - https://learn.microsoft.com/en-us/power-apps/developer/data-platform/query-antipatterns
//   - https://learn.microsoft.com/en-us/power-apps/developer/data-platform/bypass-custom-business-logic
//   - https://learn.microsoft.com/en-us/power-apps/developer/data-platform/bypass-power-automate-flows

import type { ReactNode } from 'react';

/**
 * Severity tiers — drive the iconography + counter chip color in AdvisoryDrawer.
 *
 *   error   → red — execute is blocked. Always surfaced inline at the source AND
 *                   in the drawer. The mode should pass the same message to
 *                   <UrlBar disabledReason>.
 *   warning → amber — execute proceeds but the request will likely be slow,
 *                     throttled, or have unintended side-effects (e.g. plug-in
 *                     bypass that affects business rules).
 *   info    → blue — a transparent rewrite happened (wildcard strip), a header
 *                    was auto-added, or the user picked a non-obvious option
 *                    that's worth flagging.
 */
export type AdvisorySeverity = 'error' | 'warning' | 'info';

/**
 * Source categories — used purely for counter aggregation + filtering inside
 * the drawer.
 */
export type AdvisorySource =
  | 'wildcard'         // wildcard characters (rewrites, stripped leading wildcards)
  | 'antipattern'      // query anti-patterns
  | 'bypass'           // bypass custom business logic
  | 'flow-bypass'      // bypass Power Automate flows
  | 'privilege'        // cross-cutting privilege requirements
  | 'header'           // header conflicts / auto-injection notice
  | 'validation';      // missing required field / shape error

export interface Advisory {
  /** Stable id for React keys + dedupe across re-renders. */
  id: string;
  severity: AdvisorySeverity;
  source: AdvisorySource;
  /** Short noun-phrase headline, ~50 chars. Shown collapsed in the chip count. */
  title: string;
  /** Optional longer prose body. Rendered as ReactNode so it can include <code>/<a>. */
  body?: ReactNode;
  /**
   * Which clause / sidebar node to focus when the user clicks "go fix".
   * Modes wire this to their onSelect handler.
   */
  focusNode?: string;
  /**
   * Optional MS Learn deep link — the drawer renders this as a "Learn more"
   * link below the body.
   */
  learnMoreUrl?: string;
}

// ── Helpers — short factories for the common cases ───────────────────────────

export const adv = {
  err: (id: string, source: AdvisorySource, title: string, body?: ReactNode, focusNode?: string, learnMoreUrl?: string): Advisory =>
    ({ id, severity: 'error', source, title, body, focusNode, learnMoreUrl }),
  warn: (id: string, source: AdvisorySource, title: string, body?: ReactNode, focusNode?: string, learnMoreUrl?: string): Advisory =>
    ({ id, severity: 'warning', source, title, body, focusNode, learnMoreUrl }),
  info: (id: string, source: AdvisorySource, title: string, body?: ReactNode, focusNode?: string, learnMoreUrl?: string): Advisory =>
    ({ id, severity: 'info', source, title, body, focusNode, learnMoreUrl }),
};

/**
 * Bucket advisories by severity. Used by the drawer to render in order
 * (errors first, then warnings, then info) and to drive the counter chip.
 */
export function bucketAdvisories(advisories: Advisory[]) {
  const errors:   Advisory[] = [];
  const warnings: Advisory[] = [];
  const infos:    Advisory[] = [];
  for (const a of advisories) {
    if (a.severity === 'error') errors.push(a);
    else if (a.severity === 'warning') warnings.push(a);
    else infos.push(a);
  }
  return { errors, warnings, infos, total: advisories.length };
}

/**
 * The single message that should be passed to <UrlBar disabledReason>.
 * If there are no errors, returns null. If multiple, joins with " · ".
 * The drawer is where the user reads the detail; this string is just the
 * tooltip on the disabled Execute button.
 */
export function disabledReasonFromAdvisories(advisories: Advisory[]): string | null {
  const errs = advisories.filter(a => a.severity === 'error');
  if (errs.length === 0) return null;
  if (errs.length === 1) return errs[0].title;
  return `${errs.length} blockers — open Advisories to review.`;
}
