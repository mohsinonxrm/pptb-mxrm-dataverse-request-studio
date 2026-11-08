// #293: the model-driven form URL for a returned record.
import { describe, expect, it } from 'vitest';
import { recordFormUrl } from './openInDataverse';

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
