// Bypass-related advisory generator — surfaces privilege requirements,
// legacy-header warnings, step-cap blockers, and Power-Automate "courtesy
// reminder" advisories.
//
// Called by every write mode alongside its other advisory generators so the
// shared AdvisoryDrawer renders a single aggregated list.
//
// Reference: https://learn.microsoft.com/en-us/power-apps/developer/data-platform/bypass-custom-business-logic
//            https://learn.microsoft.com/en-us/power-apps/developer/data-platform/bypass-power-automate-flows

import type { BypassOptions } from '../state/writeState';
import { adv, type Advisory } from '../primitives/advisories';
import {
  STEP_IDS_DEFAULT_LIMIT,
  STEP_IDS_MAX_LIMIT,
  PRV_BYPASS_BUSINESS_LOGIC,
  PRV_BYPASS_CUSTOM_PLUGINS,
} from './bypassHeaders';

const LEARN_BYPASS_LOGIC =
  'https://learn.microsoft.com/en-us/power-apps/developer/data-platform/bypass-custom-business-logic';
const LEARN_BYPASS_FLOWS =
  'https://learn.microsoft.com/en-us/power-apps/developer/data-platform/bypass-power-automate-flows';

export function detectBypassAdvisories(b: BypassOptions, focusNode = 'bypass'): Advisory[] {
  const out: Advisory[] = [];

  // ── Business-logic bypass ─────────────────────────────────────────
  if (b.businessLogic !== 'none') {
    const privilege =
      b.useLegacyHeader && b.businessLogic === 'sync'
        ? PRV_BYPASS_CUSTOM_PLUGINS
        : PRV_BYPASS_BUSINESS_LOGIC;

    out.push(
      adv.warn(
        'bypass-active',
        'bypass',
        `Bypassing ${labelForMode(b)} business logic`,
        <span>
          This request requires the <code>{privilege.name}</code> privilege on the caller.{' '}
          {privilege.description}
        </span>,
        focusNode,
        LEARN_BYPASS_LOGIC,
      ),
    );

    // Step-cap enforcement
    if (b.businessLogic === 'steps') {
      const count = b.stepIds.filter(Boolean).length;
      if (count === 0) {
        out.push(
          adv.err(
            'bypass-no-steps',
            'validation',
            'Step bypass enabled but no step IDs provided',
            <span>Add at least one plug-in step GUID, or change the bypass mode to "None".</span>,
            focusNode,
            LEARN_BYPASS_LOGIC,
          ),
        );
      } else if (count > STEP_IDS_MAX_LIMIT) {
        out.push(
          adv.err(
            'bypass-step-cap-max',
            'validation',
            `${count} step IDs — exceeds the maximum cap (${STEP_IDS_MAX_LIMIT})`,
            <span>Dataverse rejects more than {STEP_IDS_MAX_LIMIT} step IDs.</span>,
            focusNode,
            LEARN_BYPASS_LOGIC,
          ),
        );
      } else if (count > STEP_IDS_DEFAULT_LIMIT) {
        out.push(
          adv.warn(
            'bypass-step-cap-default',
            'bypass',
            `${count} step IDs — exceeds the default org cap (${STEP_IDS_DEFAULT_LIMIT})`,
            <span>
              Default cap is {STEP_IDS_DEFAULT_LIMIT}. Anything above this requires raising
              <code> BypassBusinessLogicExecutionStepIdsLimit </code> via the OrgDbOrgSettings tool.
            </span>,
            focusNode,
            LEARN_BYPASS_LOGIC,
          ),
        );
      }
    }

    // Legacy-header escape hatch
    if (b.useLegacyHeader && b.businessLogic === 'sync') {
      out.push(
        adv.info(
          'bypass-legacy-header',
          'bypass',
          'Using legacy MSCRM.BypassCustomPluginExecution header',
          <span>
            The newer <code>MSCRM.BypassBusinessLogicExecution: CustomSync</code> is the recommended
            replacement. Both wire effects are equivalent for sync-only bypass.
          </span>,
          focusNode,
          LEARN_BYPASS_LOGIC,
        ),
      );
    }
  }

  // ── Power Automate flow suppression ───────────────────────────────
  if (b.suppressFlows) {
    out.push(
      adv.warn(
        'bypass-flows',
        'flow-bypass',
        'Suppressing Power Automate flow triggers',
        <span>
          Flow owners are <strong>not notified</strong> when their logic is bypassed. Communicate
          with them before running this against production data — child flows may still fire via
          other triggers later.
        </span>,
        focusNode,
        LEARN_BYPASS_FLOWS,
      ),
    );
  }

  return out;
}

function labelForMode(b: BypassOptions): string {
  switch (b.businessLogic) {
    case 'sync':
      return 'synchronous';
    case 'async':
      return 'asynchronous';
    case 'both':
      return 'sync + async';
    case 'steps':
      return `${b.stepIds.filter(Boolean).length} specific plug-in step(s) of`;
    case 'none':
    default:
      return '';
  }
}
