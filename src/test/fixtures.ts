// Shared metadata fixtures for unit tests. A tiny account / contact model
// with the column shapes the encoders special-case: lookups (single-target
// and polymorphic Customer), multi-select choices, money / integer /
// boolean / datetime, an alternate key, and both nav cardinalities.

import { __clearLiveTables, __registerLiveTables, type TableMeta } from '../mock/metadata';
import {
  _setTableLookup,
  type FilterGroup,
  type FilterNode,
  type FilterRule,
} from '../editors/filter/filterTree';
import { findTable } from '../mock/metadata';
import { emptyApply } from '../editors/ApplyEditor';
import { emptyPrefer } from '../editors/PreferEditor';
import type { RetrieveMultipleState } from '../state/readState';
import { defaultBypassOptions, type CreateState } from '../state/writeState';

export const ACCOUNT_ID = '00000000-0000-0000-0000-0000000000a1';
export const CONTACT_ID = '00000000-0000-0000-0000-0000000000c1';

export const account: TableMeta = {
  logicalName: 'account',
  entitySetName: 'accounts',
  displayName: 'Account',
  primaryKey: 'accountid',
  primaryName: 'name',
  supportsMerge: true,
  columns: [
    { logicalName: 'accountid', displayName: 'Account', attributeType: 'Uniqueidentifier' },
    {
      logicalName: 'name',
      displayName: 'Name',
      attributeType: 'String',
      format: 'Text',
      maxLength: 160,
    },
    {
      logicalName: 'accountnumber',
      displayName: 'Account Number',
      attributeType: 'String',
      format: 'Text',
      maxLength: 20,
    },
    {
      logicalName: 'numberofemployees',
      displayName: 'Employees',
      attributeType: 'Integer',
      format: 'None',
    },
    {
      logicalName: 'revenue',
      displayName: 'Revenue',
      attributeType: 'Money',
      precisionSource: 2,
      precision: 2,
    },
    {
      logicalName: 'donotemail',
      displayName: 'Do not allow Emails',
      attributeType: 'Boolean',
      trueOption: { value: 1, label: 'Do Not Allow' },
      falseOption: { value: 0, label: 'Allow' },
    },
    {
      logicalName: 'createdon',
      displayName: 'Created On',
      attributeType: 'DateTime',
      format: 'DateAndTime',
      dateTimeBehavior: 'UserLocal',
    },
    {
      logicalName: 'industrycode',
      displayName: 'Industry',
      attributeType: 'Picklist',
      options: [
        { value: 1, label: 'Accounting' },
        { value: 2, label: 'Agriculture' },
      ],
    },
    {
      logicalName: 'sample_tags',
      displayName: 'Tags',
      attributeType: 'MultiSelectPicklist',
      options: [
        { value: 100, label: 'A' },
        { value: 200, label: 'B' },
      ],
    },
    {
      logicalName: 'primarycontactid',
      displayName: 'Primary Contact',
      attributeType: 'Lookup',
      targets: ['contact'],
      oDataName: '_primarycontactid_value',
    },
  ],
  navigationProperties: [
    {
      name: 'primarycontactid',
      relationshipName: 'account_primary_contact',
      cardinality: 'ManyToOne',
      targetEntity: 'contact',
      referencingAttribute: 'primarycontactid',
    },
    {
      name: 'contact_customer_accounts',
      relationshipName: 'contact_customer_accounts',
      cardinality: 'OneToMany',
      targetEntity: 'contact',
    },
  ],
  alternateKeys: [
    { name: 'akey_accountnumber', displayName: 'By Account Number', columns: ['accountnumber'] },
    {
      name: 'akey_composite',
      displayName: 'By Name + Employees',
      columns: ['name', 'numberofemployees'],
    },
  ],
};

export const contact: TableMeta = {
  logicalName: 'contact',
  entitySetName: 'contacts',
  displayName: 'Contact',
  primaryKey: 'contactid',
  primaryName: 'fullname',
  columns: [
    { logicalName: 'contactid', displayName: 'Contact', attributeType: 'Uniqueidentifier' },
    {
      logicalName: 'fullname',
      displayName: 'Full Name',
      attributeType: 'String',
      format: 'Text',
      maxLength: 160,
    },
    {
      logicalName: 'jobtitle',
      displayName: 'Job Title',
      attributeType: 'String',
      format: 'Text',
      maxLength: 100,
    },
    {
      logicalName: 'parentcustomerid',
      displayName: 'Company Name',
      attributeType: 'Customer',
      targets: ['account', 'contact'],
      oDataName: '_parentcustomerid_value',
    },
  ],
  navigationProperties: [
    {
      name: 'parentcustomerid_account',
      relationshipName: 'contact_customer_accounts',
      cardinality: 'ManyToOne',
      targetEntity: 'account',
      referencingAttribute: 'parentcustomerid',
    },
    {
      name: 'parentcustomerid_contact',
      relationshipName: 'contact_customer_contacts',
      cardinality: 'ManyToOne',
      targetEntity: 'contact',
      referencingAttribute: 'parentcustomerid',
    },
  ],
};

/** Register the fixture tables in the live-metadata registry (what the
 *  encoders read via `findTable`) and wire the filter encoder's lambda
 *  table lookup the same way App.tsx does. */
export function registerFixtureTables(): void {
  __clearLiveTables();
  __registerLiveTables([account, contact]);
  _setTableLookup(findTable);
}

let nodeSeq = 0;
const id = (p: string) => `${p}_${++nodeSeq}`;

export function group(
  combinator: 'and' | 'or',
  rules: FilterNode[],
  extra: Partial<FilterGroup> = {},
): FilterGroup {
  return { id: id('g'), type: 'group', combinator, rules, ...extra };
}

export function rule(
  col: string,
  op: string,
  val?: string,
  extra: Partial<FilterRule> = {},
): FilterRule {
  return { id: id('r'), type: 'rule', col, op, val, ...extra };
}

export function retrieveMultipleState(
  overrides: Partial<RetrieveMultipleState> = {},
): RetrieveMultipleState {
  return {
    table: 'account',
    select: [],
    filter: group('and', []),
    orderby: [],
    top: null,
    countOn: false,
    expand: [],
    apply: emptyApply(),
    prefer: emptyPrefer(),
    headers: [],
    dirty: new Set(),
    ...overrides,
  };
}

export function createState(overrides: Partial<CreateState> = {}): CreateState {
  return {
    table: 'account',
    fieldValues: {},
    nullFields: [],
    prefer: emptyPrefer(),
    headers: [],
    returnSelect: [],
    duplicateDetection: false,
    bypass: defaultBypassOptions(),
    dirty: new Set(),
    ...overrides,
  };
}
