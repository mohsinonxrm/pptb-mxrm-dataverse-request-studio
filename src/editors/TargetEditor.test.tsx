// TargetEditor without an entity scope (#42): the shared empty state, a
// disabled picker, and Choose scope opening the header's scope popover.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';

const host = vi.hoisted(() => ({
  needsSetup: true,
  entities: [] as Array<{ logicalName: string; displayName: string; entitySetName: string }>,
  openScope: vi.fn(),
}));
vi.mock('../host/useScopedEntities', () => ({
  useScopedEntities: () => ({
    entities: host.entities,
    allEntities: host.entities,
    allLoading: false,
    loading: false,
    error: null,
    retry: () => {},
    needsSetup: host.needsSetup,
    scopeMode: 'publisher-solution',
  }),
}));
vi.mock('../host/useLiveMetadata', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../host/useLiveMetadata')>()),
  useLiveTable: () => ({ table: undefined, loading: false, error: null }),
}));
vi.mock('../host/useOpenScope', () => ({ useOpenScope: () => host.openScope }));

import { TargetEditor } from './TargetEditor';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
// Fluent's MessageBar measures its layout; jsdom has no ResizeObserver.
globalThis.ResizeObserver ??= class {
  observe() {}
  unobserve() {}
  disconnect() {}
} as unknown as typeof ResizeObserver;

let root: Root | null = null;
let el: HTMLDivElement | null = null;

function render(table = '') {
  act(() =>
    root!.render(
      <FluentProvider theme={webLightTheme}>
        <TargetEditor table={table} onTableChange={() => {}} />
      </FluentProvider>,
    ),
  );
}
const picker = () => el!.querySelector<HTMLInputElement>('input[role=combobox]')!;

beforeEach(() => {
  host.needsSetup = true;
  host.entities = [];
  host.openScope.mockReset();
  el = document.createElement('div');
  document.body.appendChild(el);
  root = createRoot(el);
});
afterEach(() => {
  act(() => root?.unmount());
  el?.remove();
  root = null;
});

describe('TargetEditor without a scope', () => {
  it('shows the shared empty state and disables the picker', () => {
    render();
    expect(el!.textContent).toContain('No entities listed yet');
    expect(el!.textContent).toContain('or switch to All entities to see everything.');
    expect(el!.textContent).not.toContain('Open Settings');
    expect(picker().disabled).toBe(true);
    expect(picker().placeholder).toBe('Choose a scope to list entities');

    const choose = [...el!.querySelectorAll('button')].find(
      (b) => b.textContent?.trim() === 'Choose scope',
    );
    act(() => choose!.click());
    expect(host.openScope).toHaveBeenCalledTimes(1);
  });

  it('keeps a table already picked clearable', () => {
    render('account');
    expect(el!.textContent).toContain('No entities listed yet');
    expect(picker().disabled).toBe(false);
  });

  it('with a scope, counts the entities in it', () => {
    host.needsSetup = false;
    host.entities = [
      { logicalName: 'account', displayName: 'Account', entitySetName: 'accounts' },
      { logicalName: 'contact', displayName: 'Contact', entitySetName: 'contacts' },
    ];
    render();
    expect(el!.textContent).not.toContain('No entities listed yet');
    expect(picker().disabled).toBe(false);
    expect(el!.textContent).toContain('2 entities in scope');
  });
});
