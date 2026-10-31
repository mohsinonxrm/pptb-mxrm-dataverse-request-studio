// Settings → Request history (#226): turning it off clears what was kept,
// and Clear history asks first.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';

const confirm = vi.hoisted(() => ({ answer: true, asked: 0 }));
vi.mock('./dialogs', async (importOriginal) => ({
  ...(await importOriginal<typeof import('./dialogs')>()),
  useConfirm: () => async () => {
    confirm.asked++;
    return confirm.answer;
  },
}));

import { SettingsDrawer } from './SettingsDrawer';
import { defaultDisplaySettings, type DisplaySettings } from '../state/displaySettings';
import { __resetHistoryCache, readHistory, recordRuns } from '../state/runHistory';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let root: Root | null = null;
let host: HTMLDivElement | null = null;
let changed: DisplaySettings | null = null;

beforeEach(() => {
  __resetHistoryCache();
  localStorage.clear();
  confirm.asked = 0;
  changed = null;
  host = document.createElement('div');
  document.body.appendChild(host);
  root = createRoot(host);
  act(() =>
    root!.render(
      <FluentProvider theme={webLightTheme}>
        <SettingsDrawer
          open
          settings={defaultDisplaySettings}
          onClose={() => {}}
          onSettingsChange={(s) => (changed = s)}
        />
      </FluentProvider>,
    ),
  );
  recordRuns(
    [{ id: 'r1', modeId: 'create', url: '/x', method: 'POST', ts: 1, status: 201, ms: 1 }],
    'org1',
  );
});
afterEach(() => {
  act(() => root?.unmount());
  host?.remove();
});

const button = (text: string) =>
  [...document.querySelectorAll('button')].find((b) => b.textContent?.trim() === text)!;

describe('Settings → Request history (#226)', () => {
  it('turning it off clears the kept runs', () => {
    const toggle = document.querySelector<HTMLInputElement>(
      'input[aria-label="Keep request history"]',
    )!;
    expect(toggle.checked).toBe(true);
    act(() => toggle.click());
    expect(changed?.keepRunHistory).toBe(false);
    expect(readHistory('org1')).toEqual([]);
  });

  it('Clear history asks first, and clears only when confirmed', async () => {
    confirm.answer = false;
    await act(async () => button('Clear history').click());
    expect(confirm.asked).toBe(1);
    expect(readHistory('org1')).toHaveLength(1);
    confirm.answer = true;
    await act(async () => button('Clear history').click());
    expect(readHistory('org1')).toEqual([]);
  });
});
