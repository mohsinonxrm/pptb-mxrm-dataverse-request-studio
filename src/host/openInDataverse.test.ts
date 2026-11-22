// #293: the model-driven form URL for a returned record.
import { describe, expect, it } from 'vitest';
import { formTableFor, recordFormUrl } from './openInDataverse';

const ID = '8038A91F-99C9-F111-AAAF-000D3A5B6760';

describe('recordFormUrl', () => {
  it('builds the entityrecord URL', () => {
    expect(recordFormUrl('org.crm.dynamics.com', 'Account', `{${ID}}`)).toBe(
      `https://org.crm.dynamics.com/main.aspx?pagetype=entityrecord&etn=account&id=${ID.toLowerCase()}`,
    );
  });

  it('is null without an environment, table or GUID', () => {
    expect(recordFormUrl('', 'account', ID)).toBeNull();
    expect(recordFormUrl('org.crm.dynamics.com', '', ID)).toBeNull();
    expect(recordFormUrl('org.crm.dynamics.com', 'account', 'not-a-guid')).toBeNull();
    expect(recordFormUrl('org.crm.dynamics.com', 'account', 42)).toBeNull();
  });
});

describe('formTableFor (#422)', () => {
  it('is the table itself for a table with a form', () => {
    expect(formTableFor('Account')).toBe('account');
    expect(formTableFor('account', { name: 'x' }, false)).toBe('account');
  });

  it('is nothing for an intersect table or no table', () => {
    expect(formTableFor('accountleads', {}, true)).toBeUndefined();
    expect(formTableFor(undefined)).toBeUndefined();
    expect(formTableFor('  ')).toBeUndefined();
  });

  it('opens an activity as its own type, and not at all without it', () => {
    expect(formTableFor('activitypointer', { activitytypecode: 'email' })).toBe('email');
    expect(formTableFor('activitypointer', { activitytypecode: 'msdyn_ocliveworkitem' })).toBe(
      'msdyn_ocliveworkitem',
    );
    expect(formTableFor('activitypointer', {})).toBeUndefined();
    expect(formTableFor('activitypointer')).toBeUndefined();
    expect(formTableFor('activitypointer', { activitytypecode: 'x&etn=y' })).toBeUndefined();
  });
});
