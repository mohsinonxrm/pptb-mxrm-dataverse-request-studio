// The Prefer pane's Track changes switch (#233): offered only when the table
// has change tracking enabled, and off with the reason otherwise.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { PreferEditor, emptyPrefer, type PreferSpec } from './PreferEditor';
import { changeTrackingAvailability } from '../engine/changeTracking';
import { account } from '../test/fixtures';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let root: Root | null = null;
let el: HTMLDivElement | null = null;

function render(
  spec: PreferSpec,
  setSpec: (s: PreferSpec) => void,
  changeTrackingEnabled?: boolean,
) {
  act(() =>
    root!.render(
      <FluentProvider theme={webLightTheme}>
        <PreferEditor
          spec={spec}
          setSpec={setSpec}
          sendsPageSize
          changeTracking={changeTrackingAvailability({
            table: 'account',
            tableMeta: { ...account, changeTrackingEnabled },
          })}
        />
      </FluentProvider>,
    ),
  );
}
const trackSwitch = () => {
  const label = [...el!.querySelectorAll('label')].find((l) => l.textContent === 'Track changes');
  return el!.querySelector<HTMLInputElement>(`#${CSS.escape(label!.htmlFor)}`)!;
};

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

describe('PreferEditor Track changes (#233)', () => {
  it('is disabled with the reason when ChangeTrackingEnabled is false', () => {
    render(emptyPrefer(), () => {}, false);
    expect(trackSwitch().disabled).toBe(true);
    expect(el!.textContent).toContain("Change tracking isn't enabled for this table.");
  });

  it('turns on for a tracked table', () => {
    const setSpec = vi.fn();
    render(emptyPrefer(), setSpec, true);
    expect(trackSwitch().disabled).toBe(false);
    expect(el!.textContent).not.toContain("isn't enabled");
    act(() => trackSwitch().click());
    expect(setSpec).toHaveBeenCalledWith({ ...emptyPrefer(), trackChanges: true });
  });

  it('can still be turned off when the table no longer allows it', () => {
    const setSpec = vi.fn();
    render({ ...emptyPrefer(), trackChanges: true }, setSpec, false);
    expect(trackSwitch().disabled).toBe(false);
    act(() => trackSwitch().click());
    expect(setSpec).toHaveBeenCalledWith({ ...emptyPrefer(), trackChanges: undefined });
  });
});
