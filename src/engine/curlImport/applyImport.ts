// Fold an imported request into a mode's fresh state (#271). Each mode passes
// its own initial state, so defaults it owns (built-in header rows, hints)
// stay as they are and only what the command carried changes.

import type { HeaderItem } from '../../editors/HeadersEditor';
import { emptyApply } from '../../editors/ApplyEditor';
import type { RetrieveMultipleState, RetrieveSingleState } from '../../state/readState';
import type { CreateState, DeleteState, UpdateState, UpsertState } from '../../state/writeState';
import type { ExecuteActionState } from '../../state/executeState';
import type { BatchState } from '../batch/batchModel';
import type { ImportedHeader, ImportedRequest } from './curlImport';

type Imported<K extends ImportedRequest['modeId']> = Extract<ImportedRequest, { modeId: K }>;

let seq = 0;

/**
 * Set the command's headers on the mode's rows: a row with the same name
 * (any case) takes the value and is turned on; any other header is added.
 */
export function mergeHeaders(base: HeaderItem[], extra: ImportedHeader[]): HeaderItem[] {
  const out = base.map((h) => ({ ...h }));
  for (const { name, value } of extra) {
    const row = out.find((h) => h.name.toLowerCase() === name.toLowerCase());
    if (row) {
      row.value = value;
      row.enabled = true;
    } else {
      out.push({
        id: `imp${Date.now().toString(36)}${(seq++).toString(36)}`,
        name,
        value,
        enabled: true,
      });
    }
  }
  return out;
}

export function importRetrieveMultiple(
  base: RetrieveMultipleState,
  r: Imported<'retrieve-multiple'>,
): RetrieveMultipleState {
  const p = r.parsed;
  return {
    ...base,
    table: p.table,
    select: p.select,
    filter: p.filter,
    orderby: p.orderby,
    top: p.top,
    countOn: p.countOn,
    countOnly: p.countOnly ? true : undefined,
    expand: p.expand,
    apply: p.apply ?? emptyApply(),
    prefer: r.prefer,
    headers: mergeHeaders(base.headers, r.headers),
    dirty: new Set(['target']),
  };
}

export function importRetrieveSingle(
  base: RetrieveSingleState,
  r: Imported<'retrieve-single'>,
): RetrieveSingleState {
  return {
    ...base,
    table: r.table,
    recordId: r.recordId,
    select: r.select,
    expand: r.expand,
    prefer: r.prefer,
    headers: mergeHeaders(base.headers, r.headers),
    dirty: new Set(['target']),
  };
}

export function importCreate(base: CreateState, r: Imported<'create'>): CreateState {
  const headers = mergeHeaders(base.headers, r.headers);
  return {
    ...base,
    table: r.table,
    fieldValues: {},
    nullFields: [],
    rawBody: r.rawBody,
    prefer: r.prefer,
    returnSelect: r.returnSelect,
    bypass: r.bypass,
    headers,
    duplicateDetection: headers.some(
      (h) =>
        h.enabled &&
        h.name.toLowerCase() === 'mscrm.suppressduplicatedetection' &&
        h.value.trim().toLowerCase() === 'false',
    ),
    dirty: new Set(['target']),
  };
}

export function importUpdate(base: UpdateState, r: Imported<'update'>): UpdateState {
  return {
    ...base,
    table: r.table,
    recordId: r.recordId,
    method: 'PATCH',
    putColumn: null,
    fieldValues: {},
    nullFields: [],
    rawBody: r.rawBody,
    prefer: r.prefer,
    returnSelect: r.returnSelect,
    concurrency: r.concurrency,
    bypass: r.bypass,
    headers: mergeHeaders(base.headers, r.headers),
    dirty: new Set(['target']),
  };
}

export function importUpsert(base: UpsertState, r: Imported<'upsert'>): UpsertState {
  return {
    ...base,
    table: r.table,
    key: { kind: 'guid', recordId: r.recordId },
    fieldValues: {},
    nullFields: [],
    rawBody: r.rawBody,
    prefer: r.prefer,
    returnSelect: r.returnSelect,
    concurrency: r.concurrency,
    bypass: r.bypass,
    headers: mergeHeaders(base.headers, r.headers),
    dirty: new Set(['target']),
  };
}

export function importDelete(base: DeleteState, r: Imported<'delete'>): DeleteState {
  return {
    ...base,
    table: r.table,
    recordId: r.recordId,
    scope: r.scope,
    concurrency: r.concurrency,
    bypassCustomPlugins: false,
    bypass: r.bypass,
    headers: mergeHeaders(base.headers, r.headers),
    confirmText: '',
    acknowledged: false,
    dirty: new Set(['target']),
  };
}

export function importExecuteAction(
  base: ExecuteActionState,
  r: Imported<'exec-action' | 'exec-customapi' | 'exec-customaction'>,
): ExecuteActionState {
  return {
    ...base,
    actionName: r.actionName,
    boundRecordId: r.boundRecordId,
    paramValues: r.paramValues,
    headers: mergeHeaders(base.headers, r.headers),
    dirty: new Set(['pick']),
  };
}

export function importBatch(base: BatchState, r: Imported<'batch'>): BatchState {
  return {
    ...base,
    groups: r.groups,
    continueOnError: r.continueOnError,
    dirty: new Set(['overview']),
  };
}
