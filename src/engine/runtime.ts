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
import { invalidateLookupRecords } from '../host/useLookupRecords';

// Writes go through `afterWrite`, which drops cached record-picker searches
// once the request settles, so pickers re-query instead of showing a
// record's old name (#81). It runs even when the call fails, since a
// failed or timed-out response doesn't prove nothing was written.
// `table` is the table written to; `undefined` drops every table — a delete
// can cascade, a create can deep-insert related rows, and actions and
// workflows can touch anything.
// Trailing arguments (the write group's request headers, #43) pass through
// unchanged.
function afterWrite<S, Rest extends unknown[]>(
  run: (s: S, ...rest: Rest) => Promise<ExecResult>,
  table: (s: S) => string | undefined,
): (s: S, ...rest: Rest) => Promise<ExecResult> {
  return async (s, ...rest) => {
    try {
      return await run(s, ...rest);
    } finally {
      invalidateLookupRecords(table(s));
    }
  };
}
const everyTable = () => undefined;
const ownTable = (s: { table: string }) => s.table;

// One dispatch wrapper per executor method.
export const runtime = neverThrow({
  // Reads, relate and Execute Function: the mode's composed headers; the
  // executor sends an allow-list of them (#109).
  retrieveMultiple: (s: RetrieveMultipleState, headers?: Record<string, string>) =>
    liveRetrieveMultiple(s, headers),
  retrieveSingle: (s: RetrieveSingleState, headers?: Record<string, string>) =>
    liveRetrieveSingle(s, headers),
  nextLink: (s: RetrieveNextLinkState, headers?: Record<string, string>) =>
    liveNextLink(s, headers),
  absoluteUrl: (url: string, maxpagesize?: number | null, headers?: Record<string, string>) =>
    liveAbsoluteUrl(url, maxpagesize, headers),
  predefined: (s: PredefinedQueryState, headers?: Record<string, string>) =>
    livePredefined(s, headers),
  create: afterWrite(
    (s: CreateState, headers?: Record<string, string>) => liveCreate(s, headers),
    everyTable,
  ),
  update: afterWrite(
    (s: UpdateState, headers?: Record<string, string>) => liveUpdate(s, headers),
    ownTable,
  ),
  upsert: afterWrite(
    (s: UpsertState, headers?: Record<string, string>) => liveUpsert(s, headers),
    ownTable,
  ),
  delete: afterWrite(
    (s: DeleteState, headers?: Record<string, string>) => liveDelete(s, headers),
    everyTable,
  ),
  merge: afterWrite(
    (s: MergeState, headers?: Record<string, string>) => liveMerge(s, headers),
    ownTable,
  ),
  associate: afterWrite(
    (s: AssociateState, headers?: Record<string, string>) => liveAssociate(s, headers),
    ownTable,
  ),
  disassociate: afterWrite(
    (s: DisassociateState, headers?: Record<string, string>) => liveDisassociate(s, headers),
    ownTable,
  ),
  action: afterWrite(
    (s: ExecuteActionState, headers?: Record<string, string>) => liveAction(s, headers),
    everyTable,
  ),
  // Functions are side-effect free by definition — nothing to invalidate.
  function: (s: ExecuteFunctionState, headers?: Record<string, string>) => liveFunction(s, headers),
  workflow: afterWrite(
    (s: ExecuteWorkflowState, headers?: Record<string, string>) => liveWorkflow(s, headers),
    everyTable,
  ),
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
