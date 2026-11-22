// A Batch / ExecuteMultiple request without an entity scope (#420): its table
// and URL suggestions list the scope's tables, so one line says why there are
// none yet, and its link opens the scope picker.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';

const host = vi.hoisted(() => ({ needsSetup: true, openScope: vi.fn() }));
vi.mock('../../host/useScopedEntities', () => ({
  useScopedEntities: () => ({
    entities: [],
    allEntities: [],
    allLoading: false,
    loading: false,
    error: null,
    retry: () => {},
    needsSetup: host.needsSetup,
    scopeMode: 'publisher-solution',
  }),
}));
vi.mock('../../host/useOpenScope', () => ({ useOpenScope: () => host.openScope }));

import { BatchItemEditor } from './BatchItemEditor';
import { newGroup, newItem } from '../../engine/batch/batchModel';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
globalThis.ResizeObserver ??= class {
  observe() {}
  unobserve() {}
  disconnect() {}
} as unknown as typeof ResizeObserver;

let root: Root | null = null;
let el: HTMLDivElement | null = null;

function render() {
  const item = newItem('retrieve');
  const group = newGroup(false, [item]);
  act(() =>
    root!.render(
      <FluentProvider theme={webLightTheme}>
        <BatchItemEditor
          item={item}
          contentId="1"
          group={group}
          groups={[group]}
          wire={null}
          issues={[]}
          onChange={() => {}}
          onRemove={() => {}}
          onMove={() => {}}
          onMoveToGroup={() => {}}
        />
      </FluentProvider>,
    ),
  );
}

beforeEach(() => {
  host.needsSetup = true;
  host.openScope.mockReset();
  el = document.createElement('div');
  document.body.appendChild(el);
  root = createRoot(el);
});
afterEach(() => {
  act(() => root?.unmount());
  el?.remove();
});

describe('BatchItemEditor without a scope (#420)', () => {
  it('says tables are suggested once a scope is chosen, and the link opens it', () => {
    render();
    expect(el!.textContent).toContain('Tables are suggested once you choose a scope.');
    // One line with a link, not another primary button or a second scope card.
    expect(el!.querySelector('.fui-Button--primary, button[class*="primary"]')).toBeNull();
    expect(el!.textContent).not.toContain('Step 1');
    const link = [...el!.querySelectorAll('button, a')].find(
      (b) => b.textContent === 'choose a scope',
    ) as HTMLElement;
    act(() => link.click());
    expect(host.openScope).toHaveBeenCalledTimes(1);
  });

  it('says nothing with a scope', () => {
    host.needsSetup = false;
    render();
    expect(el!.textContent).not.toContain('choose a scope');
  });
});
