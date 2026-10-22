import {
  Table20Regular,
  Document20Regular,
  Filter20Regular,
  ChevronRight20Regular,
  Edit20Regular,
  Delete20Regular,
  DeleteLines20Regular,
  BranchFork20Regular,
  Link20Regular,
  Flash20Regular,
  Code20Regular,
  Flowchart20Regular,
  Image20Regular,
  Attach20Regular,
  Stack20Regular,
} from '@fluentui/react-icons';
import type { FC } from 'react';

export type RequestGroup = 'read' | 'write' | 'relate' | 'execute' | 'binary' | 'bulk' | 'batch';
export type HttpMethod = 'GET' | 'POST' | 'PATCH' | 'PUT' | 'DELETE' | 'BATCH';

export interface RequestType {
  id: string;
  name: string;
  /**
   * Primary HTTP method — shown on the request-type chip in the header.
   * For cardinality-dependent modes (Associate, Disassociate), this is the
   * collection-valued verb; `altMethod` carries the single-valued verb so
   * the chip can render both ("POST/PATCH"). The actual verb on the wire
   * is computed by the mode at runtime.
   */
  method: HttpMethod;
  /**
   * Optional secondary HTTP method — used by the request-type chip to show
   * dual-verb modes (Associate: POST collection / PATCH single, Disassociate:
   * DELETE collection / PATCH single). When set, MethodPill renders both
   * verbs joined with a slash so the user understands the mode's verb is
   * cardinality-dependent.
   */
  altMethod?: HttpMethod;
  group: RequestGroup;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  icon: FC<any>;
  sub: string;
  // True when DRS has an actual mode component wired up for this id.
  // Modes set to `false` route to StubMode.
  implemented: boolean;
  // Verb shown on the Execute button (e.g. "Execute", "Send", "Run").
  executeVerb: string;
}

export const REQ_TYPES: RequestType[] = [
  // Read
  {
    id: 'retrieve-multiple',
    name: 'Retrieve Multiple',
    method: 'GET',
    group: 'read',
    icon: Table20Regular,
    sub: 'Query a table',
    implemented: true,
    executeVerb: 'Execute',
  },
  {
    id: 'retrieve-single',
    name: 'Retrieve Single',
    method: 'GET',
    group: 'read',
    icon: Document20Regular,
    sub: 'By ID',
    implemented: true,
    executeVerb: 'Execute',
  },
  {
    id: 'retrieve-nextlink',
    name: 'Retrieve NextLink',
    method: 'GET',
    group: 'read',
    icon: ChevronRight20Regular,
    sub: 'Pagination',
    implemented: true,
    executeVerb: 'Execute',
  },
  {
    id: 'predefined-query',
    name: 'Predefined Query',
    method: 'GET',
    group: 'read',
    icon: Filter20Regular,
    sub: 'Saved or user query',
    implemented: true,
    executeVerb: 'Execute',
  },

  // Write
  {
    id: 'create',
    name: 'Create Record',
    method: 'POST',
    group: 'write',
    icon: Document20Regular,
    sub: 'Insert one',
    implemented: true,
    executeVerb: 'Send',
  },
  {
    id: 'update',
    name: 'Update Record',
    method: 'PATCH',
    group: 'write',
    icon: Edit20Regular,
    sub: 'Modify fields',
    implemented: true,
    executeVerb: 'Send',
  },
  {
    id: 'upsert',
    name: 'Upsert Record',
    method: 'PATCH',
    group: 'write',
    icon: Edit20Regular,
    sub: 'Create or update',
    implemented: true,
    executeVerb: 'Send',
  },
  {
    id: 'delete',
    name: 'Delete Record',
    method: 'DELETE',
    group: 'write',
    icon: Delete20Regular,
    sub: 'Remove one',
    implemented: true,
    executeVerb: 'Send',
  },
  // Bulk — BulkDelete asynchronous system job from a query (#10).
  {
    id: 'bulk-delete',
    name: 'Bulk Delete Job',
    method: 'POST',
    group: 'bulk',
    icon: DeleteLines20Regular,
    sub: 'Async BulkDelete from a query',
    implemented: true,
    executeVerb: 'Send',
  },
  {
    id: 'merge',
    name: 'Merge Records',
    method: 'POST',
    group: 'write',
    icon: BranchFork20Regular,
    sub: 'Two records into one',
    implemented: true,
    executeVerb: 'Send',
  },
  // Bulk — CreateMultiple / UpdateMultiple / UpsertMultiple / DeleteMultiple
  // (collection-bound actions with a Targets array, #8).
  {
    id: 'bulk-write',
    name: 'Bulk Write',
    method: 'POST',
    group: 'bulk',
    icon: Stack20Regular,
    sub: 'Create/Update/Upsert/DeleteMultiple',
    implemented: true,
    executeVerb: 'Send',
  },

  // Relate
  // Associate / Disassociate are cardinality-dependent — verb on the wire
  // is POST/$ref (collection) or PATCH @odata.bind (single-valued) for
  // Associate; DELETE/$ref (collection) or PATCH @odata.bind: null
  // (single-valued) for Disassociate. The dual-verb chip in the picker
  // reflects this so the user knows the verb isn't fixed.
  {
    id: 'associate',
    name: 'Associate',
    method: 'POST',
    altMethod: 'PATCH',
    group: 'relate',
    icon: Link20Regular,
    sub: 'Link records',
    implemented: true,
    executeVerb: 'Send',
  },
  {
    id: 'disassociate',
    name: 'Disassociate',
    method: 'DELETE',
    altMethod: 'PATCH',
    group: 'relate',
    icon: Link20Regular,
    sub: 'Unlink records',
    implemented: true,
    executeVerb: 'Send',
  },

  // Execute
  {
    id: 'exec-action',
    name: 'Execute Action',
    method: 'POST',
    group: 'execute',
    icon: Flash20Regular,
    sub: 'Bound or unbound',
    implemented: true,
    executeVerb: 'Run',
  },
  {
    id: 'exec-function',
    name: 'Execute Function',
    method: 'GET',
    group: 'execute',
    icon: Flash20Regular,
    sub: 'OData function',
    implemented: true,
    executeVerb: 'Run',
  },
  {
    id: 'exec-customapi',
    name: 'Execute Custom API',
    method: 'POST',
    group: 'execute',
    icon: Code20Regular,
    sub: 'Custom API',
    implemented: true,
    executeVerb: 'Run',
  },
  {
    id: 'exec-customaction',
    name: 'Execute Custom Action',
    method: 'POST',
    group: 'execute',
    icon: Code20Regular,
    sub: 'Custom Action',
    implemented: true,
    executeVerb: 'Run',
  },
  {
    id: 'exec-workflow',
    name: 'Execute Workflow',
    method: 'POST',
    group: 'execute',
    icon: Flowchart20Regular,
    sub: 'On-demand workflow',
    implemented: true,
    executeVerb: 'Run',
  },

  // Binary
  {
    id: 'manage-file',
    name: 'Manage File Data',
    method: 'PATCH',
    group: 'binary',
    icon: Document20Regular,
    sub: 'File column',
    implemented: true,
    executeVerb: 'Run',
  },
  {
    id: 'manage-image',
    name: 'Manage Image Data',
    method: 'PATCH',
    group: 'binary',
    icon: Image20Regular,
    sub: 'Image column',
    implemented: true,
    executeVerb: 'Run',
  },
  {
    id: 'manage-attachment',
    name: 'Manage Attachment / Note',
    method: 'POST',
    group: 'binary',
    icon: Attach20Regular,
    sub: 'Attachment & Annotation',
    implemented: true,
    executeVerb: 'Run',
  },

  // Batch
  {
    id: 'batch',
    name: '$batch Request',
    method: 'BATCH',
    group: 'batch',
    icon: Stack20Regular,
    sub: 'Change sets & multiple requests',
    implemented: true,
    executeVerb: 'Send',
  },
  {
    id: 'execute-multiple',
    name: 'ExecuteMultiple',
    method: 'BATCH',
    group: 'batch',
    icon: Stack20Regular,
    sub: 'SDK-style bulk requests via $batch',
    implemented: true,
    executeVerb: 'Send',
  },
];

export const GROUPS: { id: RequestGroup; label: string }[] = [
  { id: 'read', label: 'Read' },
  { id: 'write', label: 'Write' },
  { id: 'relate', label: 'Relate' },
  { id: 'execute', label: 'Execute' },
  { id: 'batch', label: 'Batch' },
  { id: 'binary', label: 'Binary' },
  { id: 'bulk', label: 'Bulk' },
];

export const findRequestType = (id: string): RequestType =>
  REQ_TYPES.find((r) => r.id === id) ?? REQ_TYPES[0];
