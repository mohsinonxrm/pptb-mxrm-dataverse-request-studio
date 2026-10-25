import { describe, expect, it } from 'vitest';
import { writeResultName } from './writeResultName';

describe('writeResultName (#209)', () => {
  it('uses the returned primary name, e.g. a computed contact fullname', () => {
    expect(
      writeResultName(
        'fullname',
        { fullname: 'DRS E2E Retrieve', contactid: 'x' },
        { lastname: 'Retrieve' },
      ),
    ).toBe('DRS E2E Retrieve');
  });

  it('falls back to the value that was sent', () => {
    expect(writeResultName('name', null, { name: 'Contoso' })).toBe('Contoso');
    expect(writeResultName('name', { name: '' }, { name: 'Contoso' })).toBe('Contoso');
  });

  it('is null when neither has a name, or the table is unknown', () => {
    expect(writeResultName('fullname', {}, { lastname: 'x' })).toBeNull();
    expect(writeResultName(undefined, { fullname: 'x' }, {})).toBeNull();
  });
});
