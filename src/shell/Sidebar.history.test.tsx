// Sidebar Recent runs with a kept history (#226): restoring a run whose
// caller id the history removed says so.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';

const notices = vi.hoisted(() => [] as string[]);
vi.mock('./dialogs', async (importOriginal) => ({
  ...(await importOriginal<typeof import('./dialogs')>()),
  notify: (message: string) => notices.push(message),
}));

import { Sidebar } from './Sidebar';
import { findRequestType } from '../registry/requestTypes';
import { __resetHistoryCache } from '../state/runHistory';
import type { RecentRun } from '../state/readState';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let root: Root | null = null;
let host: HTMLDivElement | null = null;
beforeEach(() => {
  __resetHistoryCache();
  localStorage.clear();
  notices.length = 0;
  host = document.createElement('div');
  document.body.appendChild(host);
  root = createRoot(host);
});
afterEach(() => {
  act(() => root?.unmount());
  host?.remove();
});

const run = (impersonationCleared?: boolean): RecentRun => ({
  id: 'r-1',
  modeId: 'update',
  url: '/api/data/v9.2/accounts(1)',
  method: 'PATCH',
  ts: Date.now(),
  status: 204,
  ms: 5,
  snapshot: { table: 'account' },
  impersonationCleared,
});

function render(r: RecentRun, onRecentClick: (r: RecentRun) => void) {
  act(() =>
    root!.render(
      <FluentProvider theme={webLightTheme}>
        <Sidebar
          type={findRequestType('update')}
          urlPreview=""
          sections={[]}
          activeNode=""
          onSelect={() => {}}
          recents={[r]}
          onRecentClick={onRecentClick}
        />
      </FluentProvider>,
    ),
  );
}
const row = () => host!.querySelector<HTMLElement>('[aria-label^="Restore run"]')!;

describe('Sidebar restore with history (#226)', () => {
  it('says the caller id was not restored', () => {
    const onRecentClick = vi.fn();
    render(run(true), onRecentClick);
    act(() => row().click());
    expect(onRecentClick).toHaveBeenCalledOnce();
    expect(notices[0]).toMatch(/runs as you/);
  });

  it('says nothing for a run without impersonation', () => {
    render(run(), vi.fn());
    act(() => row().click());
    expect(notices).toEqual([]);
  });
});
