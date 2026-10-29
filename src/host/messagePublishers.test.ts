// #248: who published a Custom API or Custom Action.
import { describe, expect, it } from 'vitest';
import {
  buildPublisherIndex,
  isMicrosoftPublisher,
  resolvePublisher,
  type SolutionPublisherRow,
} from './messagePublishers';

// Shapes seen on a Dynamics org (CRM279420), plus a customer's and an ISV's.
const ROWS: SolutionPublisherRow[] = [
  {
    solutionid: 'AAAAAAAA-0000-0000-0000-000000000001',
    uniquename: 'msdyn_SalesCopilot',
    publisherid: {
      uniquename: 'microsoftdynamics',
      friendlyname: 'Dynamics 365',
      customizationprefix: 'msdyn',
    },
  },
  {
    solutionid: 'aaaaaaaa-0000-0000-0000-000000000002',
    uniquename: 'System',
    publisherid: {
      uniquename: 'MicrosoftCorporation',
      friendlyname: 'MicrosoftCorporation',
      customizationprefix: '',
    },
  },
  {
    solutionid: 'aaaaaaaa-0000-0000-0000-000000000003',
    uniquename: 'Active',
    publisherid: {
      uniquename: 'DefaultPublisherorg1',
      friendlyname: 'Default Publisher for org1',
      customizationprefix: 'new',
    },
  },
  {
    solutionid: 'aaaaaaaa-0000-0000-0000-000000000004',
    uniquename: 'ContosoCore',
    publisherid: { uniquename: 'contoso', friendlyname: 'Contoso', customizationprefix: 'contoso' },
  },
  {
    solutionid: 'aaaaaaaa-0000-0000-0000-000000000005',
    uniquename: 'FabrikamTools',
    publisherid: {
      uniquename: 'fabrikam',
      friendlyname: 'Fabrikam ISV',
      customizationprefix: 'fab',
    },
  },
  // A solution whose publisher couldn't be expanded is skipped.
  { solutionid: 'aaaaaaaa-0000-0000-0000-000000000006', uniquename: 'Orphan', publisherid: null },
];

const index = buildPublisherIndex(ROWS);
const SOL = (n: number) => `aaaaaaaa-0000-0000-0000-00000000000${n}`;

describe('isMicrosoftPublisher', () => {
  it('knows Microsoft’s publishers', () => {
    for (const uniquename of [
      'MicrosoftCorporation',
      'microsoftfirstparty',
      'microsoftdynamics',
      'dynamics365customerengagement',
      'DynamicsMKT',
      'microsoftformspro',
      'QMSAnchor',
      'adxstudio',
    ]) {
      expect(isMicrosoftPublisher({ uniquename }), uniquename).toBe(true);
    }
  });

  it('recognises unknown Microsoft-looking publishers by name or prefix', () => {
    expect(isMicrosoftPublisher({ uniquename: 'MicrosoftDynamicsFieldService' })).toBe(true);
    expect(
      isMicrosoftPublisher({ uniquename: 'x', friendlyname: 'Dynamics 365 Field Service' }),
    ).toBe(true);
    expect(isMicrosoftPublisher({ uniquename: 'x', friendlyname: 'Microsoft Power Pages' })).toBe(
      true,
    );
    expect(isMicrosoftPublisher({ uniquename: 'x', customizationprefix: 'msdynce' })).toBe(true);
    expect(isMicrosoftPublisher({ uniquename: 'x', customizationprefix: 'adx' })).toBe(true);
  });

  it('leaves customers and ISVs alone', () => {
    expect(
      isMicrosoftPublisher({
        uniquename: 'contoso',
        friendlyname: 'Contoso',
        customizationprefix: 'contoso',
      }),
    ).toBe(false);
    expect(
      isMicrosoftPublisher({
        uniquename: 'DefaultPublisherorg1',
        friendlyname: 'Default Publisher for org1',
        customizationprefix: 'new',
      }),
    ).toBe(false);
    expect(
      isMicrosoftPublisher({
        uniquename: 'Cr738e5',
        friendlyname: 'CDS Default Publisher',
        customizationprefix: 'cr6f5',
      }),
    ).toBe(false);
    // "Dynamics" alone isn't Microsoft: an ISV can be called that.
    expect(
      isMicrosoftPublisher({
        uniquename: 'dyncon',
        friendlyname: 'Dynamics Consulting',
        customizationprefix: 'dc',
      }),
    ).toBe(false);
    // Nor is a customer's prefix that happens to start with "ms".
    expect(
      isMicrosoftPublisher({
        uniquename: 'morganstanley',
        friendlyname: 'Morgan Stanley',
        customizationprefix: 'ms',
      }),
    ).toBe(false);
    expect(
      isMicrosoftPublisher({
        uniquename: 'mscind',
        friendlyname: 'MSC',
        customizationprefix: 'msc',
      }),
    ).toBe(false);
  });
});

describe('resolvePublisher', () => {
  it('uses the owning solution’s publisher, whatever the GUID’s case', () => {
    expect(resolvePublisher('msdyn_DoThing', SOL(1), index)).toEqual({
      name: 'Dynamics 365',
      microsoft: true,
    });
    expect(resolvePublisher('contoso_DoThing', SOL(4).toUpperCase(), index)).toEqual({
      name: 'Contoso',
      microsoft: false,
    });
    expect(resolvePublisher('fab_Sync', SOL(5), index)).toEqual({
      name: 'Fabrikam ISV',
      microsoft: false,
    });
    // The System solution is Microsoft's.
    expect(resolvePublisher('CalculateRollupField', SOL(2), index)).toEqual({
      name: 'MicrosoftCorporation',
      microsoft: true,
    });
  });

  it('reads an unmanaged layer (Active solution) by the name’s customization prefix', () => {
    expect(resolvePublisher('new_MyAction', SOL(3), index)).toEqual({
      name: 'Default Publisher for org1',
      microsoft: false,
    });
    expect(resolvePublisher('contoso_Draft', SOL(3), index)).toEqual({
      name: 'Contoso',
      microsoft: false,
    });
    // A managed Microsoft action seen with an Active solutionid (ClassifySegments, live).
    expect(resolvePublisher('ClassifySegments', SOL(3), index)).toEqual({
      name: 'Microsoft',
      microsoft: true,
    });
  });

  it('falls back to the prefix when the solution is unknown', () => {
    expect(resolvePublisher('msdyn_Other', undefined, index)).toEqual({
      name: 'Dynamics 365',
      microsoft: true,
    });
    expect(resolvePublisher('acme_Run', 'ffffffff-0000-0000-0000-000000000000', index)).toEqual({
      name: 'acme',
      microsoft: false,
    });
    expect(resolvePublisher('mspcat_Run', undefined, index)).toEqual({
      name: 'mspcat',
      microsoft: true,
    });
    expect(resolvePublisher('acme_Run', SOL(6), index)).toEqual({ name: 'acme', microsoft: false });
  });
});
