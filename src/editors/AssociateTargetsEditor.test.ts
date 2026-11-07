// #317: the targets title (pane head and sidebar, Associate and Disassociate)
// is singular for a single-valued relationship or one queued target.
import { describe, expect, it } from 'vitest';
import { targetsTitle } from './AssociateTargetsEditor';

describe('targetsTitle', () => {
  it('is singular for a single-valued relationship, whatever is queued', () => {
    expect(targetsTitle(true, 0)).toBe('Target record');
    expect(targetsTitle(true, 1)).toBe('Target record');
  });

  it('follows the number of targets for a collection', () => {
    expect(targetsTitle(false, 0)).toBe('Target records');
    expect(targetsTitle(false, 1)).toBe('Target record');
    expect(targetsTitle(false, 3)).toBe('Target records');
  });
});
