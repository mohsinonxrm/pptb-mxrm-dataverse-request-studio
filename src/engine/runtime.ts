// Executor runtime — single dispatch point for live execution.
//
// PPTB-only. Every `runtime.X(state)` call routes to the matching live
// executor in `dataverseExecutor.ts`. When called outside PPTB the
// `dvHost.*` helpers throw `HostNotAvailableError` — that's the correct
// behavior for a PPTB-targeted product.

import {
  executeRetrieveMultiple as liveRetrieveMultiple,
  executeRetrieveSingle   as liveRetrieveSingle,
  executeNextLink         as liveNextLink,
  executeAbsoluteUrl      as liveAbsoluteUrl,
  executePredefined       as livePredefined,
  executeCreate           as liveCreate,
  executeUpdate           as liveUpdate,
  executeUpsert           as liveUpsert,
  executeDelete           as liveDelete,
  executeMerge            as liveMerge,
  executeAssociate        as liveAssociate,
  executeDisassociate     as liveDisassociate,
  executeAction           as liveAction,
  executeFunction         as liveFunction,
  executeWorkflow         as liveWorkflow,
  executeManageFile       as liveManageFile,
  executeManageImage      as liveManageImage,
  executeManageAttachment as liveManageAttachment,
  type ExecResult,
} from './dataverseExecutor';

import type { RetrieveMultipleState, RetrieveSingleState, RetrieveNextLinkState, PredefinedQueryState } from '../state/readState';
import type { CreateState, UpdateState, UpsertState, DeleteState, MergeState } from '../state/writeState';
import type { AssociateState, DisassociateState } from '../state/relateState';
import type { ExecuteActionState, ExecuteFunctionState, ExecuteWorkflowState } from '../state/executeState';
import type { ManageFileState, ManageImageState, ManageAttachmentState } from '../state/binaryState';
import { isEmbedded } from '../host/pptbBridge';

// One dispatch wrapper per executor method.
export const runtime = {
  retrieveMultiple: (s: RetrieveMultipleState) => liveRetrieveMultiple(s),
  retrieveSingle:   (s: RetrieveSingleState)   => liveRetrieveSingle(s),
  nextLink:         (s: RetrieveNextLinkState) => liveNextLink(s),
  absoluteUrl:      (url: string)              => liveAbsoluteUrl(url),
  predefined:       (s: PredefinedQueryState)  => livePredefined(s),
  create:           (s: CreateState)           => liveCreate(s),
  update:           (s: UpdateState)           => liveUpdate(s),
  upsert:           (s: UpsertState)           => liveUpsert(s),
  delete:           (s: DeleteState)           => liveDelete(s),
  merge:            (s: MergeState)            => liveMerge(s),
  associate:        (s: AssociateState)        => liveAssociate(s),
  disassociate:     (s: DisassociateState)     => liveDisassociate(s),
  action:           (s: ExecuteActionState)    => liveAction(s),
  function:         (s: ExecuteFunctionState)  => liveFunction(s),
  workflow:         (s: ExecuteWorkflowState)  => liveWorkflow(s),
  manageFile:       (s: ManageFileState)       => liveManageFile(s),
  manageImage:      (s: ManageImageState)      => liveManageImage(s),
  manageAttachment: (s: ManageAttachmentState) => liveManageAttachment(s),
};

export type { ExecResult };

/**
 * Whether the runtime is connected to a real Dataverse host. True when
 * embedded in PPTB. Modes use this to gate execute-time affordances
 * (e.g. surface a "running standalone — Execute will fail" hint).
 */
export const isLiveRuntime = (): boolean => isEmbedded();
