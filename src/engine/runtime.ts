// Executor runtime — single dispatch point for live execution.
//
// PPTB-only. Every `runtime.X(state)` call routes to the matching live
// executor in `dataverseExecutor.ts`. When called outside PPTB the
// `dvHost.*` helpers throw `HostNotAvailableError` — that's the correct
// behavior for a PPTB-targeted product. No dispatch rejects: a throw comes
// back as a failed `ExecResult` (see `neverThrow`).

import {
  executeRetrieveMultiple as liveRetrieveMultiple,
  executeRetrieveSingle as liveRetrieveSingle,
  executeNextLink as liveNextLink,
  executeAbsoluteUrl as liveAbsoluteUrl,
  executePredefined as livePredefined,
  executeCreate as liveCreate,
  executeUpdate as liveUpdate,
  executeUpsert as liveUpsert,
  executeDelete as liveDelete,
  executeMerge as liveMerge,
  executeAssociate as liveAssociate,
  executeDisassociate as liveDisassociate,
  executeAction as liveAction,
  executeFunction as liveFunction,
  executeWorkflow as liveWorkflow,
  executeManageFile as liveManageFile,
  executeManageImage as liveManageImage,
  executeManageAttachment as liveManageAttachment,
  execFailure,
  type ExecResult,
} from './dataverseExecutor';

import type {
  RetrieveMultipleState,
  RetrieveSingleState,
  RetrieveNextLinkState,
  PredefinedQueryState,
} from '../state/readState';
import type {
  CreateState,
  UpdateState,
  UpsertState,
  DeleteState,
  MergeState,
} from '../state/writeState';
import type { AssociateState, DisassociateState } from '../state/relateState';
import type {
  ExecuteActionState,
  ExecuteFunctionState,
  ExecuteWorkflowState,
} from '../state/executeState';
import type {
  ManageFileState,
  ManageImageState,
  ManageAttachmentState,
} from '../state/binaryState';
import { isEmbedded } from '../host/pptbBridge';

// One dispatch wrapper per executor method.
export const runtime = neverThrow({
  retrieveMultiple: (s: RetrieveMultipleState) => liveRetrieveMultiple(s),
  retrieveSingle: (s: RetrieveSingleState) => liveRetrieveSingle(s),
  nextLink: (s: RetrieveNextLinkState) => liveNextLink(s),
  absoluteUrl: (url: string) => liveAbsoluteUrl(url),
  predefined: (s: PredefinedQueryState) => livePredefined(s),
  create: (s: CreateState, headers?: Record<string, string>) => liveCreate(s, headers),
  update: (s: UpdateState, headers?: Record<string, string>) => liveUpdate(s, headers),
  upsert: (s: UpsertState, headers?: Record<string, string>) => liveUpsert(s, headers),
  delete: (s: DeleteState, headers?: Record<string, string>) => liveDelete(s, headers),
  merge: (s: MergeState, headers?: Record<string, string>) => liveMerge(s, headers),
  associate: (s: AssociateState) => liveAssociate(s),
  disassociate: (s: DisassociateState) => liveDisassociate(s),
  action: (s: ExecuteActionState) => liveAction(s),
  function: (s: ExecuteFunctionState) => liveFunction(s),
  workflow: (s: ExecuteWorkflowState) => liveWorkflow(s),
  manageFile: (s: ManageFileState) => liveManageFile(s),
  manageImage: (s: ManageImageState) => liveManageImage(s),
  manageAttachment: (s: ManageAttachmentState) => liveManageAttachment(s),
});

export type { ExecResult };

/**
 * Whether the runtime is connected to a real Dataverse host. True when
 * embedded in PPTB. Modes use this to gate execute-time affordances
 * (e.g. surface a "running standalone — Execute will fail" hint).
 */
export const isLiveRuntime = (): boolean => isEmbedded();

/**
 * Turns anything a dispatch throws into a failed `ExecResult` (same mapping
 * as a host error). Modes `await runtime.X(...)` and then clear their
 * spinner, so a throw there (e.g. `$metadata` failing to load while
 * resolving an action) used to leave Execute spinning and the rejection
 * unhandled (#102).
 */
function neverThrow<T extends Record<string, (...args: never[]) => Promise<ExecResult>>>(
  dispatch: T,
): T {
  const guarded: Record<string, (...args: unknown[]) => Promise<ExecResult>> = {};
  for (const [name, run] of Object.entries(dispatch)) {
    const call = run as unknown as (...args: unknown[]) => Promise<ExecResult>;
    guarded[name] = async (...args) => {
      const t0 = performance.now();
      try {
        return await call(...args);
      } catch (e) {
        return execFailure(e, Math.round(performance.now() - t0));
      }
    };
  }
  return guarded as unknown as T;
}
