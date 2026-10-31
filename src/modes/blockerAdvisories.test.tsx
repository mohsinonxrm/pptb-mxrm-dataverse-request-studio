// Every mode reports its Execute blocker as an advisory that leads to the
// fix (#294): the blocker is an error advisory, the Execute tooltip names it,
// its focusNode is a node of that mode's sidebar, and Open selects that node.
import { act, type ComponentType } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import type { Advisory } from '../primitives/advisories';

interface SidebarNode {
  id: string;
  children?: SidebarNode[];
}
const host = vi.hoisted(() => ({
  urlBar: null as null | {
    advisories?: Advisory[];
    disabledReason?: string | null;
    onAdvisoryFocus?: (id: string) => void;
  },
  sidebar: null as null | {
    sections: { id: string; items: SidebarNode[] }[];
    activeNode: string;
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

import { CreateMode } from './CreateMode';
import { AssociateMode } from './AssociateMode';
import { DisassociateMode } from './DisassociateMode';
import { ExecuteActionMode } from './ExecuteActionMode';
import { ExecuteFunctionMode } from './ExecuteFunctionMode';
import { ExecuteWorkflowMode } from './ExecuteWorkflowMode';
import { ManageFileMode } from './ManageFileMode';
import { ManageImageMode } from './ManageImageMode';
import { ManageAttachmentMode } from './ManageAttachmentMode';
import { PredefinedQueryMode } from './PredefinedQueryMode';
import { BatchMode } from './BatchMode';
import { BulkWriteMode } from './BulkWriteMode';
import { BulkDeleteMode } from './BulkDeleteMode';
import { SaveContextRoot, type SaveContextValue } from '../state/SaveContext';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
globalThis.ResizeObserver ??= class {
  observe() {}
  unobserve() {}
  disconnect() {}
} as unknown as typeof ResizeObserver;

let root: Root | null = null;
let hostEl: HTMLDivElement | null = null;
let published: SaveContextValue | null = null;

beforeEach(() => {
  host.urlBar = null;
  host.sidebar = null;
  published = null;
  hostEl = document.createElement('div');
  document.body.appendChild(hostEl);
  root = createRoot(hostEl);
});
afterEach(() => {
  act(() => root?.unmount());
  hostEl?.remove();
  root = null;
});

function render(Mode: ComponentType) {
  act(() =>
    root!.render(
      <FluentProvider theme={webLightTheme}>
        <SaveContextRoot value={published} setValue={(v) => (published = v)}>
          <Mode />
        </SaveContextRoot>
      </FluentProvider>,
    ),
  );
}

const nodeIds = (nodes: SidebarNode[]): string[] =>
  nodes.flatMap((n) => [n.id, ...nodeIds(n.children ?? [])]);

// A fresh mode's first blocker, and the node that fixes it.
const MODES: [string, ComponentType, string, string][] = [
  ['Create', () => <CreateMode themeMode="light" />, 'create-table', 'target'],
  ['Associate', () => <AssociateMode themeMode="light" />, 'assoc-table', 'source'],
  ['Disassociate', () => <DisassociateMode themeMode="light" />, 'disassoc-table', 'source'],
  [
    'Execute Action',
    () => <ExecuteActionMode themeMode="light" category="oob" />,
    'action-pick',
    'pick',
  ],
  ['Execute Function', () => <ExecuteFunctionMode themeMode="light" />, 'function-pick', 'pick'],
  ['Execute Workflow', () => <ExecuteWorkflowMode themeMode="light" />, 'workflow-pick', 'pick'],
  ['Manage File', () => <ManageFileMode themeMode="light" />, 'file-table', 'target'],
  ['Manage Image', () => <ManageImageMode themeMode="light" />, 'image-table', 'target'],
  [
    'Manage Attachment',
    () => <ManageAttachmentMode themeMode="light" />,
    'attach-parent',
    'parent',
  ],
  ['Predefined Query', () => <PredefinedQueryMode themeMode="light" />, 'pq-table', 'target'],
  // Outside ToolBox a batch can't be sent at all; the overview says why.
  ['Batch', () => <BatchMode themeMode="light" />, 'batch-unavailable', 'overview'],
  ['Bulk Write', () => <BulkWriteMode themeMode="light" />, 'bulk-table', 'target'],
  ['Bulk Delete', () => <BulkDeleteMode themeMode="light" />, 'bd-table', 'target'],
];

describe('Execute blockers as advisories (#294)', () => {
  it.each(MODES)('%s: the blocker leads to its sidebar node', (_, Mode, id, focusNode) => {
    render(Mode);
    const errors = (host.urlBar?.advisories ?? []).filter((a) => a.severity === 'error');
    expect(errors.map((a) => a.id)).toContain(id);
    const blocker = errors.find((a) => a.id === id)!;
    expect(blocker.focusNode).toBe(focusNode);
    // The Execute tooltip names the blocker (or counts them).
    expect(host.urlBar?.disabledReason).toBe(
      errors.length === 1
        ? blocker.title
        : `${errors.length} blockers — open Advisories to review.`,
    );
    // The node exists in this mode's sidebar, and Open selects it.
    const ids = nodeIds(host.sidebar!.sections.flatMap((s) => s.items));
    expect(ids).toContain(focusNode);
    act(() => host.urlBar!.onAdvisoryFocus!(focusNode));
    render(Mode);
    expect(host.sidebar!.activeNode).toBe(focusNode);
  });
});
