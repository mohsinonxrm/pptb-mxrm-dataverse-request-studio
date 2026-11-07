// #297: the Builder's first-step empty state names the mode and the first step.
import { act } from 'react';
import { createRoot } from 'react-dom/client';
import { describe, expect, it, vi } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { FirstStep, ModeIdContext } from './FirstStep';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

describe('FirstStep', () => {
  it('describes the routed mode and starts the first step', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const root = createRoot(host);
    const onStart = vi.fn();
    act(() =>
      root.render(
        <FluentProvider theme={webLightTheme}>
          <ModeIdContext.Provider value="create">
            <FirstStep step="Choose a table" onStart={onStart} />
          </ModeIdContext.Provider>
        </FluentProvider>,
      ),
    );
    const region = host.querySelector('section')!;
    expect(region.getAttribute('aria-label')).toBe('Get started with Create Record');
    expect(region.textContent).toContain('Insert one.');
    expect(region.textContent).toContain('Step 1 · Choose a table');
    act(() => region.querySelector('button')!.click());
    expect(onStart).toHaveBeenCalledTimes(1);
    act(() => root.unmount());
    host.remove();
  });
});
