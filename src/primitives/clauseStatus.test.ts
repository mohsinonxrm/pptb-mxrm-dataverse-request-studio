// #417: the one badge rule (#296) for record-id clauses, as every mode uses
// it: Retrieve Single, Delete, Update, Merge (target and subordinate),
// Manage File, Manage Image, Execute Workflow and bound Execute
// Action / Function.
import { describe, expect, it } from 'vitest';
import { recordStatus, requiredStatus } from './clauseStatus';

const GUID = '8038a91f-99c9-f111-aaaf-000d3a5b6760';

describe('recordStatus (#417)', () => {
  it.each([
    // id, required, status
    [GUID, true, 'set'],
    [GUID, false, 'set'],
    // A bad id from a saved request, draft or import: the drawer lists it as a blocker.
    ['not-a-guid', true, 'invalid'],
    ['not-a-guid', false, 'invalid'],
    [`{${GUID}}`, true, 'invalid'],
    [`${GUID})/other_nav`, true, 'invalid'],
    // Missing and needed (a table is picked, a binding needs a record).
    [null, true, 'required'],
    ['', true, 'required'],
    [undefined, true, 'required'],
    // Missing, and nothing asks for it yet.
    [null, false, undefined],
    ['', false, undefined],
  ] as const)('%s (required: %s) → %s', (id, required, status) => {
    expect(recordStatus(id, required)).toBe(status);
  });
});

describe('requiredStatus (#417)', () => {
  it.each([
    // Parent rows over required parameters / fields follow their children:
    // no "N req unset" badge.
    [0, undefined],
    [1, 'required'],
    [3, 'required'],
  ] as const)('%s missing → %s', (missing, status) => {
    expect(requiredStatus(missing)).toBe(status);
  });
});
