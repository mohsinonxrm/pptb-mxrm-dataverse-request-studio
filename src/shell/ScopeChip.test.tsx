// The scope chip (#42): the first-run tip, the label for each scope, the
// popover's edits (applied right away, per org) and what each mode fetches.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';

const host = vi.hoisted(() => ({
  org: 'org1.crm.dynamics.com',
  publisherEnabled: [] as boolean[],
  solutionEnabled: [] as boolean[],
}));
vi.mock('../host/activeOrg', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../host/activeOrg')>()),
  useActiveOrg: () => host.org,
}));
const contoso = {
  publisherid: 'p1',
  friendlyname: 'Contoso',
  uniquename: 'contoso',
  customizationprefix: 'cso',
};
const fabrikam = {
  publisherid: 'p2',
  friendlyname: 'Fabrikam',
  uniquename: 'fabrikam',
  customizationprefix: 'fab',
};
const core = {
  solutionid: 's1',
  friendlyname: 'Sales Core',
  uniquename: 'cso_salescore',
  version: '1.2.0.0',
  ismanaged: false,
  _publisherid_value: 'p1',
};
const fabSol = {
  solutionid: 's2',
  friendlyname: 'Fabrikam Base',
  uniquename: 'fab_base',
  version: '3.0.0.0',
  ismanaged: true,
  _publisherid_value: 'p2',
};
vi.mock('../host/usePublisherFilter', () => ({
  usePublisherFilter: (enabled: boolean) => {
    host.publisherEnabled.push(enabled);
    const publishersWithSolutions = enabled
      ? [
          { publisher: contoso, solutions: [core] },
          { publisher: fabrikam, solutions: [fabSol] },
        ]
      : [];
    return {
      publishersWithSolutions,
      publishers: publishersWithSolutions.map((p) => p.publisher),
      publishersLoading: false,
      publishersError: null,
    };
  },
}));
vi.mock('../host/useSolutionFilter', () => ({
  useSolutionFilter: (enabled: boolean) => {
    host.solutionEnabled.push(enabled);
    return {
      solutions: enabled ? [core, fabSol] : [],
      solutionsLoading: false,
      solutionsError: null,
    };
  },
}));

import { ScopeChip } from './ScopeChip';
import { SettingsProvider, usePersistedSettings } from '../host/usePersistedSettings';
import { OpenScopeProvider, useOpenScope } from '../host/useOpenScope';
import type { DisplaySettings } from '../state/displaySettings';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let root: Root | null = null;
let el: HTMLDivElement | null = null;
let current: [DisplaySettings, (s: DisplaySettings) => void] | null = null;
let openScope: (() => void) | null = null;

function Probe() {
  current = usePersistedSettings();
  openScope = useOpenScope();
  return null;
}

function render() {
  act(() =>
    root!.render(
      <FluentProvider theme={webLightTheme}>
        <SettingsProvider>
          <OpenScopeProvider>
            <Probe />
            <ScopeChip accessSummary={null} />
          </OpenScopeProvider>
        </SettingsProvider>
      </FluentProvider>,
    ),
  );
}

const settings = () => current![0];
const last = <T,>(a: T[]) => a[a.length - 1];
const update = (patch: Partial<DisplaySettings>) =>
  act(() => current![1]({ ...current![0], ...patch }));
const chip = () => document.querySelector<HTMLButtonElement>('[data-testid=scope-chip]')!;
const bodyText = () => document.body.textContent ?? '';
const button = (text: string) =>
  [...document.querySelectorAll('button')].find((b) => b.textContent?.trim() === text);
const click = (b: Element | undefined | null) =>
  act(() => {
    if (!b) throw new Error('missing element');
    (b as HTMLElement).click();
  });

beforeEach(() => {
  host.publisherEnabled = [];
  host.solutionEnabled = [];
  el = document.createElement('div');
  document.body.appendChild(el);
  root = createRoot(el);
});
afterEach(() => {
  act(() => root?.unmount());
  el?.remove();
  document.body.innerHTML = '';
  root = null;
});

describe('ScopeChip', () => {
  it('says Choose scope and shows the tip on first run; Got it dismisses it for this environment', () => {
    render();
    expect(chip().textContent).toContain('Choose scope');
    expect(chip().getAttribute('aria-label')).toBe('Choose scope');
    expect(bodyText()).toContain('Choose where entities come from');

    click(button('Got it'));
    expect(settings().scopeTipDismissed).toEqual(['org1.crm.dynamics.com']);
    render();
    expect(bodyText()).not.toContain('Choose where entities come from');
  });

  it("doesn't show the tip without a connection", () => {
    host.org = '';
    try {
      render();
      expect(chip().textContent).toContain('Choose scope');
      expect(bodyText()).not.toContain('Choose where entities come from');
    } finally {
      host.org = 'org1.crm.dynamics.com';
    }
  });

  it("the tip's Choose scope opens the popover", () => {
    render();
    const tipButtons = [...document.querySelectorAll('button')].filter(
      (b) => b.textContent?.trim() === 'Choose scope' && b !== chip(),
    );
    expect(tipButtons).toHaveLength(1);
    click(tipButtons[0]);
    expect(bodyText()).toContain('Show entities from');
    expect(bodyText()).toContain('Changes apply right away.');
    expect(settings().scopeTipDismissed).toEqual(['org1.crm.dynamics.com']);
  });

  it('labels the scope: publisher › solution, with +N', () => {
    render();
    update({ selectedPublisherIds: ['p1'], selectedSolutionIds: ['s1'] });
    expect(chip().textContent).toContain('Contoso › Sales Core v1.2.0');
    expect(chip().getAttribute('aria-label')).toBe('Entity scope: Contoso › Sales Core v1.2.0');
    expect(bodyText()).not.toContain('Choose where entities come from');

    update({ selectedPublisherIds: ['p1', 'p2'], selectedSolutionIds: ['s1', 's2'] });
    expect(chip().textContent).toContain('Contoso +1 › Sales Core v1.2.0 +1');
  });

  it('removing a publisher also removes its solutions', () => {
    render();
    update({ selectedPublisherIds: ['p1', 'p2'], selectedSolutionIds: ['s1', 's2'] });
    click(chip());
    const fabrikamTag = document.querySelector(
      '[aria-label="Selected publishers"] button[value=p2]',
    );
    expect(fabrikamTag?.textContent).toBe('Fabrikam');
    click(fabrikamTag);
    expect(settings().selectedPublisherIds).toEqual(['p1']);
    expect(settings().selectedSolutionIds).toEqual(['s1']);
  });

  it('adding a publisher keeps the solutions already picked', () => {
    render();
    update({ selectedPublisherIds: ['p1'], selectedSolutionIds: ['s1'] });
    click(chip());
    const input = document.querySelector<HTMLInputElement>('input[aria-label=Publishers]')!;
    click(input);
    const fabrikam = [...document.querySelectorAll('.fui-TagPickerList [role=option]')].find((o) =>
      o.textContent?.startsWith('Fabrikam'),
    );
    click(fabrikam);
    expect(settings().selectedPublisherIds).toEqual(['p1', 'p2']);
    expect(settings().selectedSolutionIds).toEqual(['s1']);
  });

  it('Solution only → Publisher → solution selects the picked solutions’ publishers, so they stay visible', () => {
    render();
    update({ entityScopeMode: 'solution-only', selectedSolutionIds: ['s2'] });
    click(chip());
    const ps = [...document.querySelectorAll<HTMLInputElement>('input[type=radio]')].find(
      (r) => r.value === 'publisher-solution',
    );
    click(ps);
    expect(settings().entityScopeMode).toBe('publisher-solution');
    expect(settings().selectedPublisherIds).toEqual(['p2']);
    expect(settings().selectedSolutionIds).toEqual(['s2']);
    expect(chip().textContent).toContain('Fabrikam › Fabrikam Base v3.0.0');
  });

  it('switching to All entities hides the pickers and says All entities', () => {
    render();
    click(chip());
    expect(bodyText()).toContain('Publishers');
    const all = [...document.querySelectorAll<HTMLInputElement>('input[type=radio]')].find(
      (r) => r.value === 'all',
    );
    click(all);
    expect(settings().entityScopeMode).toBe('all');
    expect(chip().textContent).toContain('All entities');
    expect(bodyText()).not.toContain('Search publishers…');
    expect(document.querySelector('input[aria-label=Solutions]')).toBeNull();
  });

  it('fetches only the list the mode needs', () => {
    render();
    expect(last(host.publisherEnabled)).toBe(true);
    expect(last(host.solutionEnabled)).toBe(false);
    update({ entityScopeMode: 'solution-only', selectedSolutionIds: ['s2'] });
    expect(last(host.publisherEnabled)).toBe(false);
    expect(last(host.solutionEnabled)).toBe(true);
    expect(chip().textContent).toContain('Fabrikam Base v3.0.0');
    update({ entityScopeMode: 'all' });
    expect(last(host.publisherEnabled)).toBe(false);
    expect(last(host.solutionEnabled)).toBe(false);
  });

  it('opens from anywhere through useOpenScope (the empty state’s Choose scope)', () => {
    render();
    click(button('Got it'));
    expect(bodyText()).not.toContain('Show entities from');
    act(() => openScope!());
    expect(bodyText()).toContain('Show entities from');
  });
});
