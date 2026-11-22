import { describe, expect, it } from 'vitest';
import { emptyPrefer, preferToHeaderString } from './PreferEditor';

describe('preferToHeaderString', () => {
  it('puts the page size last, the only order Dataverse accepts with annotations (#219)', () => {
    expect(preferToHeaderString({ ...emptyPrefer(), maxpagesize: 50 })).toBe(
      'odata.include-annotations="OData.Community.Display.V1.FormattedValue", odata.maxpagesize=50',
    );
    expect(
      preferToHeaderString({
        ...emptyPrefer(),
        allAnnotations: true,
        returnRepresentation: true,
        maxpagesize: 10,
      }),
    ).toBe('odata.include-annotations="*", return=representation, odata.maxpagesize=10');
  });

  it('puts change tracking after the annotations and before the page size (#233)', () => {
    expect(preferToHeaderString({ ...emptyPrefer(), trackChanges: true, maxpagesize: 50 })).toBe(
      'odata.include-annotations="OData.Community.Display.V1.FormattedValue", odata.track-changes, odata.maxpagesize=50',
    );
    expect(
      preferToHeaderString({ ...emptyPrefer(), formattedValues: false, trackChanges: true }),
    ).toBe('odata.track-changes');
  });

  it('is null when nothing is set', () => {
    expect(preferToHeaderString({ ...emptyPrefer(), formattedValues: false })).toBeNull();
  });
});
