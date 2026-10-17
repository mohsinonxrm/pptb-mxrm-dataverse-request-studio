// Executor runtime — single dispatch point for live execution.
//
// PPTB-only. Every `runtime.X(state)` call routes to the matching live
// executor in `dataverseExecutor.ts`. When called outside PPTB the
// `dvHost.*` helpers throw `HostNotAvailableError` — that's the correct
// behavior for a PPTB-targeted product.

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
export const runtime = {
  retrieveMultiple: (s: RetrieveMultipleState) => liveRetrieveMultiple(s),
  retrieveSingle: (s: RetrieveSingleState) => liveRetrieveSingle(s),
  nextLink: (s: RetrieveNextLinkState) => liveNextLink(s),
  absoluteUrl: (url: string) => liveAbsoluteUrl(url),
  predefined: (s: PredefinedQueryState) => livePredefined(s),
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
  associate: afterWrite((s: AssociateState) => liveAssociate(s), ownTable),
  disassociate: afterWrite((s: DisassociateState) => liveDisassociate(s), ownTable),
  action: afterWrite((s: ExecuteActionState) => liveAction(s), everyTable),
  // Functions are side-effect free by definition — nothing to invalidate.
  function: (s: ExecuteFunctionState) => liveFunction(s),
  workflow: afterWrite((s: ExecuteWorkflowState) => liveWorkflow(s), everyTable),
  manageFile: (s: ManageFileState) => liveManageFile(s),
  manageImage: (s: ManageImageState) => liveManageImage(s),
  manageAttachment: (s: ManageAttachmentState) => liveManageAttachment(s),
};

export type { ExecResult };

/**
 * Whether the runtime is connected to a real Dataverse host. True when
 * embedded in PPTB. Modes use this to gate execute-time affordances
 * (e.g. surface a "running standalone — Execute will fail" hint).
 */
export const isLiveRuntime = (): boolean => isEmbedded();
