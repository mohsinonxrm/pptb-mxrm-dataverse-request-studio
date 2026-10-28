// #306: a failed workflow load offers Retry and doesn't claim there are none.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { WorkflowPickerPane } from './ExecuteWorkflowMode';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let root: Root | null = null;
let host: HTMLDivElement | null = null;

function render(props: Partial<Parameters<typeof WorkflowPickerPane>[0]>) {
  host = document.createElement('div');
  document.body.appendChild(host);
  root = createRoot(host);
  act(() =>
    root!.render(
      <WorkflowPickerPane
        workflows={[]}
        loading={false}
        error={null}
        workflowId={null}
        setWorkflowId={() => {}}
        {...props}
      />,
    ),
  );
  return host;
}

const retryButton = (el: HTMLElement) =>
  [...el.querySelectorAll('button')].find((b) => b.textContent?.trim() === 'Retry');

afterEach(() => {
  act(() => root?.unmount());
  host?.remove();
  root = null;
  host = null;
});

describe('WorkflowPickerPane', () => {
  it('shows the error with a Retry that calls onRetry', () => {
    const onRetry = vi.fn();
    const el = render({ error: 'Dataverse queryData failed: HTTP 401', onRetry });
    expect(el.textContent).toContain("Couldn't load workflows");
    expect(el.textContent).toContain('HTTP 401');
    const retry = retryButton(el);
    expect(retry).toBeDefined();
    act(() => retry!.click());
    expect(onRetry).toHaveBeenCalledTimes(1);
  });

  it("says the workflows didn't load instead of claiming there are none", () => {
    const el = render({ error: 'HTTP 401', onRetry: () => {} });
    expect(el.querySelector('input')?.placeholder).toBe('Workflows didn’t load');
  });

  it('has no Retry when the load can never succeed (outside PPTB)', () => {
    const el = render({
      error: 'Workflows can only be loaded from a live Dataverse connection (PPTB).',
    });
    expect(retryButton(el)).toBeUndefined();
  });

  it('keeps the empty-environment text for a successful empty load', () => {
    const el = render({});
    expect(el.querySelector('input')?.placeholder).toBe(
      'No activated workflows in this environment',
    );
  });

  it('hides the error while a retry is loading', () => {
    const el = render({ error: 'HTTP 401', loading: true, onRetry: () => {} });
    expect(el.textContent).not.toContain("Couldn't load workflows");
    expect(el.querySelector('input')?.placeholder).toBe('Loading workflows…');
  });
});
