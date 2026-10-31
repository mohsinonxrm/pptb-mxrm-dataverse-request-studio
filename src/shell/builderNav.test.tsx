// An advisory's Open (#294): from any tab it shows the Builder, selects the
// node, and focuses the first field of that node's pane.
import { act, useState } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { ModeShell } from './ModeShell';
import { MainTabs, type MainTab } from './MainTabs';
import { useOpenAdvisoryNode } from './builderNav';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let root: Root | null = null;
let hostEl: HTMLDivElement | null = null;
beforeEach(() => {
  hostEl = document.createElement('div');
  document.body.appendChild(hostEl);
  root = createRoot(hostEl);
});
afterEach(() => {
  act(() => root?.unmount());
  hostEl?.remove();
  root = null;
});

function OpenButton({ onSelect }: { onSelect: (id: string) => void }) {
  const open = useOpenAdvisoryNode(onSelect);
  return (
    <button type="button" onClick={() => open?.('headers')}>
      Open
    </button>
  );
}

function Mode({ withField, marked = false }: { withField: boolean; marked?: boolean }) {
  const [tab, setTab] = useState<MainTab>('results');
  const [node, setNode] = useState('target');
  return (
    <ModeShell sidebar={<span>node:{node}</span>} urlBar={<OpenButton onSelect={setNode} />}>
      <MainTabs tab={tab} onTabChange={setTab}>
        {tab === 'builder' ? (
          <div>
            Pane {node}
            {withField && <input aria-label="Header name" />}
            {/* A pane that shares several nodes marks each one. */}
            {marked && (
              <div data-focus-node="headers">
                <input aria-label="Header value" />
              </div>
            )}
          </div>
        ) : (
          <div>Results</div>
        )}
      </MainTabs>
    </ModeShell>
  );
}

// Open focuses after two animation frames: wait until something inside the
// main pane has focus (up to 2 s, so a busy test run doesn't time it out).
async function focusMoved() {
  for (let i = 0; i < 100; i++) {
    await act(async () => await new Promise((r) => setTimeout(r, 20)));
    if (document.activeElement?.closest('main')) return;
  }
}
const openButton = () =>
  [...hostEl!.querySelectorAll('button')].find((b) => b.textContent === 'Open')!;

describe('useOpenAdvisoryNode (#294)', () => {
  it('shows the Builder, selects the node and focuses its first field', async () => {
    act(() =>
      root!.render(
        <FluentProvider theme={webLightTheme}>
          <Mode withField />
        </FluentProvider>,
      ),
    );
    expect(hostEl!.textContent).toContain('Results');
    act(() => openButton().click());
    await focusMoved();
    expect(hostEl!.textContent).toContain('Pane headers');
    expect(hostEl!.textContent).toContain('node:headers');
    expect(document.activeElement?.getAttribute('aria-label')).toBe('Header name');
  });

  it("focuses the field of the node's own element when the pane marks it", async () => {
    act(() =>
      root!.render(
        <FluentProvider theme={webLightTheme}>
          <Mode withField marked />
        </FluentProvider>,
      ),
    );
    act(() => openButton().click());
    await focusMoved();
    expect(document.activeElement?.getAttribute('aria-label')).toBe('Header value');
  });

  it('focuses the pane itself when it has no field', async () => {
    act(() =>
      root!.render(
        <FluentProvider theme={webLightTheme}>
          <Mode withField={false} />
        </FluentProvider>,
      ),
    );
    act(() => openButton().click());
    await focusMoved();
    expect(document.activeElement?.hasAttribute('data-builder-body')).toBe(true);
  });

  it('returns nothing when the mode has no node handler', () => {
    function Probe() {
      return <span>{String(useOpenAdvisoryNode(undefined))}</span>;
    }
    act(() => root!.render(<Probe />));
    expect(hostEl!.textContent).toBe('undefined');
  });
});
