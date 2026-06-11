// Bypass-header composer — single source of truth for the four MSCRM.* headers
// described in:
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/bypass-custom-business-logic
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/bypass-power-automate-flows
//
// Returns the list of synthetic HeaderItems the mode should merge into its
// effective headers. Centralizing this means:
//   • Every write mode (Create / Update / Upsert / Delete / Merge) renders the
//     SAME bypass UI and emits the SAME wire format.
//   • Adding a new bypass option is a single-file change.
//   • Privilege metadata (prv* names + doc copy) lives next to the header
//     names, ready for the AdvisoryDrawer to surface a "requires privilege"
//     note for each bypass that's active.

import type { HeaderItem } from '../editors/HeadersEditor';
import type { BypassOptions } from '../state/writeState';

/** Privilege required to use a bypass header at the server. */
export interface BypassPrivilege {
  /** Schema name of the Dataverse privilege (e.g. `prvBypassCustomBusinessLogic`). */
  name: string;
  /** Stable GUID — required by the AddPrivilegesRole action. */
  id: string;
  /** Plain-English description. */
  description: string;
}

export const PRV_BYPASS_BUSINESS_LOGIC: BypassPrivilege = {
  name: 'prvBypassCustomBusinessLogic',
  id: '0ea552b0-a491-4470-9a1b-82068deccf66',
  description:
    'Required to use MSCRM.BypassBusinessLogicExecution + MSCRM.BypassBusinessLogicExecutionStepIds. By default only the System Administrator role has this privilege.',
};

export const PRV_BYPASS_CUSTOM_PLUGINS: BypassPrivilege = {
  name: 'prvBypassCustomPlugins',
  id: '148a9eaf-d0c4-4196-9852-c3a38e35f6a1',
  description:
    'Required to use the legacy MSCRM.BypassCustomPluginExecution header. By default only the System Administrator role has this privilege.',
};

/** Default Dataverse cap for BypassBusinessLogicExecutionStepIds — configurable
 *  via the OrgDbOrgSettings tool, max recommended is 10. */
export const STEP_IDS_DEFAULT_LIMIT = 3;
export const STEP_IDS_MAX_LIMIT = 10;

/** Internal id prefix so the user can spot composer-generated rows in the
 *  HeadersEditor table and the central composer can update them on each render. */
const PREFIX = '__bypass_';

/**
 * Compose the bypass-related synthetic HeaderItems. Returns an empty array
 * when no bypass is active.
 */
export function composeBypassHeaders(b: BypassOptions): HeaderItem[] {
  const out: HeaderItem[] = [];

  switch (b.businessLogic) {
    case 'sync':
      out.push(
        b.useLegacyHeader
          ? {
              id: `${PREFIX}plugin-legacy`,
              name: 'MSCRM.BypassCustomPluginExecution',
              value: 'true',
              enabled: true,
              builtin: true,
              hint: 'LEGACY — prefer MSCRM.BypassBusinessLogicExecution: CustomSync (requires prvBypassCustomBusinessLogic).',
            }
          : {
              id: `${PREFIX}bl-sync`,
              name: 'MSCRM.BypassBusinessLogicExecution',
              value: 'CustomSync',
              enabled: true,
              builtin: true,
              hint: 'Bypass custom synchronous plug-ins/workflows for this request. Requires prvBypassCustomBusinessLogic.',
            },
      );
      break;
    case 'async':
      out.push({
        id: `${PREFIX}bl-async`,
        name: 'MSCRM.BypassBusinessLogicExecution',
        value: 'CustomAsync',
        enabled: true,
        builtin: true,
        hint: 'Bypass custom asynchronous logic. Requires prvBypassCustomBusinessLogic.',
      });
      break;
    case 'both':
      out.push({
        id: `${PREFIX}bl-both`,
        name: 'MSCRM.BypassBusinessLogicExecution',
        value: 'CustomSync,CustomAsync',
        enabled: true,
        builtin: true,
        hint: 'Bypass both synchronous and asynchronous custom logic. Requires prvBypassCustomBusinessLogic.',
      });
      break;
    case 'steps': {
      // Step IDs are comma-separated GUIDs. Server caps the count via the
      // BypassBusinessLogicExecutionStepIdsLimit OrgDbOrgSettings.
      const ids = b.stepIds.filter(Boolean).join(',');
      if (ids) {
        out.push({
          id: `${PREFIX}bl-steps`,
          name: 'MSCRM.BypassBusinessLogicExecutionStepIds',
          value: ids,
          enabled: true,
          builtin: true,
          hint: `Bypass these specific plug-in step registrations. Default cap is ${STEP_IDS_DEFAULT_LIMIT} (max ${STEP_IDS_MAX_LIMIT}). Requires prvBypassCustomBusinessLogic.`,
        });
      }
      break;
    }
    case 'none':
    default:
      break;
  }

  if (b.suppressFlows) {
    out.push({
      id: `${PREFIX}flows`,
      name: 'MSCRM.SuppressCallbackRegistrationExpanderJob',
      value: 'true',
      enabled: true,
      builtin: true,
      hint: 'Skip Power Automate flow triggers from Dataverse events. No privilege required, but flow owners are not notified that their logic was bypassed.',
    });
  }

  return out;
}

/**
 * Apply composer output to a user's HeaderItem list — replaces any existing
 * synthetic rows (those with id starting with `__bypass_`) with the new ones,
 * preserving the user's hand-added headers in place.
 */
export function applyBypassToHeaders(
  userHeaders: HeaderItem[],
  bypass: BypassOptions,
): HeaderItem[] {
  const cleaned = userHeaders.filter((h) => !h.id.startsWith(PREFIX));
  const synth = composeBypassHeaders(bypass);
  return [...cleaned, ...synth];
}
