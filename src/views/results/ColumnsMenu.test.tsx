// The Columns menu (#292, #419): its label, and what it saves per table.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { ColumnsMenu } from './ColumnsMenu';
import type { ColumnChoice } from './exportResults';
import { SettingsProvider, usePersistedSettings } from '../../host/usePersistedSettings';
import type { DisplaySettings } from '../../state/displaySettings';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let root: Root | null = null;
let el: HTMLDivElement | null = null;
let settings: [DisplaySettings, (s: DisplaySettings) => void] | null = null;
function Probe() {
  settings = usePersistedSettings();
  return null;
}

const choice = (key: string, shown = true): ColumnChoice => ({ key, label: key, shown });

function render(choices: ColumnChoice[], table: string | undefined = 'account') {
  act(() =>
    root!.render(
      <FluentProvider theme={webLightTheme}>
        <SettingsProvider>
          <Probe />
          <ColumnsMenu table={table} choices={choices} />
        </SettingsProvider>
      </FluentProvider>,
    ),
  );
}
const trigger = () =>
  [...el!.querySelectorAll('button')].find((b) => b.textContent?.startsWith('Columns'))!;
const openMenu = () => act(() => trigger().click());
const item = (text: string) =>
  [...document.body.querySelectorAll<HTMLElement>('[role=menuitemcheckbox], [role=menuitem]')].find(
    (m) => m.textContent?.trim() === text,
  )!;
const hidden = () => settings![0].hiddenColumns;

beforeEach(() => {
  el = document.createElement('div');
  document.body.appendChild(el);
  root = createRoot(el);
});
afterEach(() => {
  act(() => root?.unmount());
  el?.remove();
  root = null;
});

describe('ColumnsMenu (#292)', () => {
  it('reads "Columns" with all shown, and "Columns · N of M" with some hidden', () => {
    render([choice('name'), choice('revenue'), choice('statecode')]);
    expect(trigger().textContent).toBe('Columns');
    render([choice('name'), choice('revenue', false), choice('statecode')]);
    expect(trigger().textContent).toBe('Columns · 2 of 3');
  });

  it('saves the unticked column for the table, keeping ones this response lacks', () => {
    render([choice('name'), choice('revenue'), choice('statecode')]);
    act(() =>
      settings![1]({ ...settings![0], hiddenColumns: { account: ['elsewhere'], contact: ['x'] } }),
    );
    openMenu();
    act(() => item('revenue').click());
    expect(hidden().account).toEqual(['elsewhere', 'revenue']);
    expect(hidden().contact).toEqual(['x']);
  });

  it('Show all columns clears the table', () => {
    render([choice('name'), choice('revenue', false)]);
    act(() => settings![1]({ ...settings![0], hiddenColumns: { account: ['revenue'] } }));
    openMenu();
    act(() => item('Show all columns').click());
    expect(hidden().account).toBeUndefined();
  });

  it("won't clear the last shown column", () => {
    render([choice('name'), choice('revenue', false)]);
    openMenu();
    expect(item('name').getAttribute('aria-disabled')).toBe('true');
  });
});
