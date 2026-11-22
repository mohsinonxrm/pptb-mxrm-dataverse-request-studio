// The one sidebar badge rule (#296) for clauses that hold a record id (#417).
// Every mode that picks a record uses it, so the same state looks the same
// everywhere and agrees with the advisory drawer ("No record selected",
// "Invalid record id").

import type { ClauseStatus } from './ClauseTreeItem';
import { isRecordGuid } from './guid';

/**
 * A record-id clause's status: `set` for a valid GUID, `invalid` for any
 * other value (a typo, or a bad id from a saved request, draft or import),
 * `required` when it's empty but needed, and nothing when it's empty and not
 * needed yet (no table picked, say).
 */
export function recordStatus(
  id: string | null | undefined,
  required: boolean,
): ClauseStatus | undefined {
  if (id) return isRecordGuid(id) ? 'set' : 'invalid';
  return required ? 'required' : undefined;
}

/**
 * A parent clause over required children (parameters, fields): `required`
 * while any of them is missing, like the children's own badges. Otherwise
 * nothing, so the clause's count badge shows.
 */
export function requiredStatus(missing: number): ClauseStatus | undefined {
  return missing > 0 ? 'required' : undefined;
}
