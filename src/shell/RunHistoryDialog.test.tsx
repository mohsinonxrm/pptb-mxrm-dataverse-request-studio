// The History dialog (#226): search, restore, and runs that can't be restored.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { RunHistoryDialog, filterRuns } from './RunHistoryDialog';
import type { RecentRun } from '../state/readState';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const run = (n: number, over: Partial<RecentRun> = {}): RecentRun => ({
  id: `r-${n}`,
  modeId: 'retrieve-multiple',
  url: `/api/data/v9.2/${n % 2 ? 'accounts' : 'contacts'}?$top=${n}`,
  method: 'GET',
  ts: Date.now() - n * 60_000,
  status: n === 3 ? 404 : 200,
  ms: 10,
  rowCount: n,
  snapshot: { top: n },
  ...over,
});

describe('filterRuns', () => {
  const runs = [run(1), run(2), run(3)];
  it('matches every word against method, URL, status and summary', () => {
    expect(filterRuns(runs, 'accounts').map((r) => r.id)).toEqual(['r-1', 'r-3']);
    expect(filterRuns(runs, 'accounts 404').map((r) => r.id)).toEqual(['r-3']);
    expect(filterRuns(runs, 'GET contacts').map((r) => r.id)).toEqual(['r-2']);
    expect(filterRuns(runs, '  ')).toHaveLength(3);
  });
});

describe('RunHistoryDialog', () => {
  let root: Root | null = null;
  let host: HTMLDivElement | null = null;
  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
    root = createRoot(host);
  });
  afterEach(() => {
    act(() => root?.unmount());
    host?.remove();
  });

  it('restores a run on click and closes; a run without a snapshot is not a button', () => {
    const onRestore = vi.fn();
    const onOpenChange = vi.fn();
    act(() =>
      root!.render(
        <FluentProvider theme={webLightTheme}>
          <RunHistoryDialog
            open
            onOpenChange={onOpenChange}
            modeName="Retrieve Multiple"
            runs={[run(1), run(2, { snapshot: undefined })]}
            onRestore={onRestore}
          />
        </FluentProvider>,
      ),
    );
    const items = [...document.querySelectorAll<HTMLElement>('[role=listitem]')];
    expect(items).toHaveLength(2);
    expect(items[0].getAttribute('aria-label')).toMatch(/^Restore run: GET/);
    expect(items[1].hasAttribute('aria-label')).toBe(false);
    act(() => items[1].click());
    expect(onRestore).not.toHaveBeenCalled();
    act(() => items[0].click());
    expect(onRestore).toHaveBeenCalledWith(expect.objectContaining({ id: 'r-1' }));
    expect(onOpenChange).toHaveBeenCalledWith(false);
  });
});
