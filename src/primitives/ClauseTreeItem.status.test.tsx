// #296: one badge rule for required, invalid and set.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, describe, expect, it } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { Table20Regular } from '@fluentui/react-icons';
import { ClauseTreeItem, type ClauseStatus } from './ClauseTreeItem';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let root: Root | null = null;
let host: HTMLDivElement | null = null;
afterEach(() => {
  act(() => root?.unmount());
  host?.remove();
});

function render(status: ClauseStatus | undefined, badge?: string) {
  host = document.createElement('div');
  document.body.appendChild(host);
  root = createRoot(host);
  act(() =>
    root!.render(
      <FluentProvider theme={webLightTheme}>
        <ClauseTreeItem icon={Table20Regular} label="Target" status={status} badge={badge} />
      </FluentProvider>,
    ),
  );
  return host.querySelector('button')!;
}

describe('ClauseTreeItem status (#296)', () => {
  it('required reads "required"', () => {
    expect(render('required').textContent).toBe('Targetrequired');
  });

  it('invalid reads "!" and is named for screen readers', () => {
    const b = render('invalid');
    expect(b.textContent).toBe('Target!');
    expect(b.querySelector('[role=img]')?.getAttribute('aria-label')).toBe('Invalid');
  });

  it('set is a named checkmark, and status wins over a badge', () => {
    const b = render('set', '3');
    expect(b.textContent).toBe('Target');
    expect(b.querySelector('[role=img]')?.getAttribute('aria-label')).toBe('Set');
  });

  it('without a status the badge shows as before', () => {
    expect(render(undefined, '3').textContent).toBe('Target3');
  });
});
