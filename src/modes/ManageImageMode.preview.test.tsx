// #224: Manage Image's preview pane (banner, preview tile, pipeline) has a
// sidebar entry, so it can be reopened after picking a record — and the
// tile then shows the record's initials.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import type { Advisory } from '../primitives/advisories';

interface SidebarNode {
  id: string;
  label?: unknown;
  badge?: unknown;
}
const host = vi.hoisted(() => ({
  embedded: false,
  urlBar: null as null | {
    advisories?: Advisory[];
    onAdvisoryFocus?: (id: string) => void;
  },
  sidebar: null as null | {
    sections: { id: string; items: SidebarNode[] }[];
    activeNode: string;
    onSelect: (id: string) => void;
  },
  target: null as null | {
    onTableChange: (t: string) => void;
    onRecordChange?: (id: string | null, primary?: string) => void;
  },
}));
vi.mock('../shell/UrlBar', () => ({
  UrlBar: (p: typeof host.urlBar) => {
    host.urlBar = p;
    return null;
  },
}));
vi.mock('../shell/Sidebar', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../shell/Sidebar')>()),
  Sidebar: (p: typeof host.sidebar) => {
    host.sidebar = p;
    return null;
  },
}));
vi.mock('../editors/TargetEditor', () => ({
  TargetEditor: (p: typeof host.target) => {
    host.target = p;
    return <div>target pane</div>;
  },
}));
vi.mock('../host/pptbBridge', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../host/pptbBridge')>()),
  isEmbedded: () => host.embedded,
}));
vi.mock('../host/useScopedEntities', () => ({
  useScopedEntities: () => ({
    entities: [],
    allEntities: [],
    allLoading: false,
    loading: false,
    error: null,
    retry: () => {},
    needsSetup: false,
    scopeMode: 'all',
  }),
}));
vi.mock('../views/CodeView', () => ({ CodeView: () => null }));
vi.mock('../views/ResultsView', () => ({ ResultsView: () => null }));

import { ManageImageMode } from './ManageImageMode';
import { clearModeDrafts } from '../state/modeDrafts';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
globalThis.ResizeObserver ??= class {
  observe() {}
  unobserve() {}
  disconnect() {}
} as unknown as typeof ResizeObserver;

let root: Root | null = null;
let hostEl: HTMLDivElement | null = null;

beforeEach(() => {
  clearModeDrafts();
  host.embedded = false;
  host.urlBar = null;
  host.sidebar = null;
  host.target = null;
});
afterEach(() => {
  act(() => root?.unmount());
  hostEl?.remove();
  root = null;
});

function mount() {
  hostEl = document.createElement('div');
  document.body.appendChild(hostEl);
  root = createRoot(hostEl);
  act(() =>
    root!.render(
      <FluentProvider theme={webLightTheme}>
        <ManageImageMode themeMode="light" />
      </FluentProvider>,
    ),
  );
}

const body = () => document.querySelector('[data-builder-body]');
const select = (id: string) => act(() => host.sidebar!.onSelect(id));
const previewNode = () =>
  host.sidebar!.sections.flatMap((s) => s.items).find((i) => i.id === 'preview');
/** The 144×144 tile: the caption above it ends in "· 144×144". */
const tileText = () =>
  [...body()!.querySelectorAll('span')]
    .find((c) => c.textContent?.endsWith('· 144×144'))
    ?.nextElementSibling?.textContent?.trim();

function pickRecord(name: string) {
  select('target');
  expect(body()?.textContent).toContain('target pane');
  act(() => host.target!.onTableChange('account'));
  act(() => host.target!.onRecordChange!('00000000-0000-0000-0000-000000000001', name));
}

describe('Manage Image preview pane (#224)', () => {
  it('opens on the preview, which has a sidebar entry', () => {
    mount();
    expect(host.sidebar!.activeNode).toBe('preview');
    expect(previewNode()).toBeDefined();
    expect(tileText()).toBe('?');
  });

  it('can be reopened after picking a record, and shows its initials', () => {
    mount();
    pickRecord('Contoso Ltd');
    expect(tileText()).toBeUndefined();
    // The sidebar offers the node; selecting it renders the preview.
    expect(previewNode()).toBeDefined();
    select('preview');
    expect(host.sidebar!.activeNode).toBe('preview');
    expect(tileText()).toBe('CL');
    expect(previewNode()?.badge).toBe('CL');
    expect(previewNode()?.label).toBe('Contoso Ltd');
  });

  it('falls back to a plain label for a name without initials', () => {
    mount();
    pickRecord('   ');
    expect(previewNode()?.label).toBe('Image preview');
    expect(previewNode()?.badge).toBeNull();
    pickRecord('###');
    expect(previewNode()?.label).toBe('###');
    expect(previewNode()?.badge).toBeNull();
  });

  it('is offered for every operation', () => {
    mount();
    for (const op of ['Download', 'Delete', 'Upload']) {
      select('operation');
      const card = [...body()!.querySelectorAll('button')].find((b) =>
        b.textContent?.startsWith(op),
      );
      act(() => card!.click());
      expect(previewNode()).toBeDefined();
      select('preview');
      expect(tileText()).toBe('?');
    }
  });

  it("Open on the 'run externally' advisory goes to the preview (in ToolBox)", () => {
    host.embedded = true;
    mount();
    pickRecord('Contoso Ltd');
    const advisory = host.urlBar!.advisories!.find((a) => a.id === 'binary-host-unsupported');
    expect(advisory?.focusNode).toBe('preview');
    // Open leads to a node the sidebar shows, so it's highlighted there.
    expect(previewNode()).toBeDefined();
    act(() => host.urlBar!.onAdvisoryFocus!(advisory!.focusNode!));
    expect(host.sidebar!.activeNode).toBe('preview');
    expect(tileText()).toBe('CL');
  });
});
