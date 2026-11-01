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
  /**
   * One plain sentence on what the request does, for people who don't know
   * the Web API (#289). Shown under the name in the picker and palette,
   * matched by search, and used as the mode card's tooltip.
   */
  description: string;
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
    description: 'Query a table with filters, sorting, paging and related rows.',
    implemented: true,
    executeVerb: 'Execute',
  },
  {
    id: 'retrieve-single',
    name: 'Retrieve Single',
    method: 'GET',
    group: 'read',
    icon: Document20Regular,
    description: 'Get one record by its ID, with the columns and related rows you choose.',
    implemented: true,
    executeVerb: 'Execute',
  },
  {
    id: 'retrieve-nextlink',
    name: 'Retrieve NextLink',
    method: 'GET',
    group: 'read',
    icon: ChevronRight20Regular,
    description: 'Get the next page of an earlier query from its @odata.nextLink.',
    implemented: true,
    executeVerb: 'Execute',
  },
  {
    id: 'predefined-query',
    name: 'Predefined Query',
    method: 'GET',
    group: 'read',
    icon: Filter20Regular,
    description: 'Run a system view or a personal view by its ID.',
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
    description: 'Add a record from a form or a JSON body.',
    implemented: true,
    executeVerb: 'Send',
  },
  {
    id: 'update',
    name: 'Update Record',
    method: 'PATCH',
    group: 'write',
    icon: Edit20Regular,
    description: 'Change columns on an existing record.',
    implemented: true,
    executeVerb: 'Send',
  },
  {
    id: 'upsert',
    name: 'Upsert Record',
    method: 'PATCH',
    group: 'write',
    icon: Edit20Regular,
    description: 'Update a record by its ID or an alternate key, or create it if it is missing.',
    implemented: true,
    executeVerb: 'Send',
  },
  {
    id: 'delete',
    name: 'Delete Record',
    method: 'DELETE',
    group: 'write',
    icon: Delete20Regular,
    description: 'Delete a record, or clear one of its columns.',
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
    description: 'Start a background job that deletes every record a query matches.',
    implemented: true,
    executeVerb: 'Send',
  },
  {
    id: 'merge',
    name: 'Merge Records',
    method: 'POST',
    group: 'write',
    icon: BranchFork20Regular,
    description: 'Merge a duplicate account, contact, lead or case into another record.',
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
    description: 'Create, update, upsert or delete many records in one call.',
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
    description: 'Link two records through a relationship.',
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
    description: 'Remove the link between two related records.',
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
    description: 'Call a built-in Dataverse action, such as WinOpportunity.',
    implemented: true,
    executeVerb: 'Run',
  },
  {
    id: 'exec-function',
    name: 'Execute Function',
    method: 'GET',
    group: 'execute',
    icon: Flash20Regular,
    description: 'Call a Dataverse function, such as WhoAmI.',
    implemented: true,
    executeVerb: 'Run',
  },
  {
    id: 'exec-customapi',
    name: 'Execute Custom API',
    method: 'POST',
    group: 'execute',
    icon: Code20Regular,
    description: 'Call a Custom API by its unique name.',
    implemented: true,
    executeVerb: 'Run',
  },
  {
    id: 'exec-customaction',
    name: 'Execute Custom Action',
    method: 'POST',
    group: 'execute',
    icon: Code20Regular,
    description: 'Call a custom process action (a classic workflow of type Action).',
    implemented: true,
    executeVerb: 'Run',
  },
  {
    id: 'exec-workflow',
    name: 'Execute Workflow',
    method: 'POST',
    group: 'execute',
    icon: Flowchart20Regular,
    description: 'Run an on-demand classic workflow against a record.',
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
    description: 'Upload, download or delete the file in a file column.',
    implemented: true,
    executeVerb: 'Run',
  },
  {
    id: 'manage-image',
    name: 'Manage Image Data',
    method: 'PATCH',
    group: 'binary',
    icon: Image20Regular,
    description: 'Upload, download or delete the image in an image column.',
    implemented: true,
    executeVerb: 'Run',
  },
  {
    id: 'manage-attachment',
    name: 'Manage Attachment / Note',
    method: 'POST',
    group: 'binary',
    icon: Attach20Regular,
    description: 'Upload, download or delete a note or email attachment.',
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
    description: 'Send several requests in one $batch, with change sets for all-or-nothing writes.',
    implemented: true,
    executeVerb: 'Send',
  },
  {
    id: 'execute-multiple',
    name: 'ExecuteMultiple',
    method: 'BATCH',
    group: 'batch',
    icon: Stack20Regular,
    description: 'Send many independent requests at once, like the SDK ExecuteMultiple message.',
    implemented: true,
    executeVerb: 'Send',
  },
];

// `tagline` sits beside the group header in the picker and palette (#289).
export const GROUPS: { id: RequestGroup; label: string; tagline: string }[] = [
  { id: 'read', label: 'Read', tagline: 'Get data out of Dataverse' },
  { id: 'write', label: 'Write', tagline: 'Create, change and delete records' },
  { id: 'relate', label: 'Relate', tagline: 'Link and unlink records' },
  { id: 'execute', label: 'Execute', tagline: 'Run actions, functions and workflows' },
  { id: 'batch', label: 'Batch', tagline: 'Send several requests at once' },
  { id: 'binary', label: 'Binary', tagline: 'Files, images and attachments' },
  { id: 'bulk', label: 'Bulk', tagline: 'Change many records at once' },
];

export const findRequestType = (id: string): RequestType =>
  REQ_TYPES.find((r) => r.id === id) ?? REQ_TYPES[0];
