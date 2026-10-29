// #241: writes ask before they're sent; the confirmation names the target, and
// on Production the word must be typed.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { AppDialogHost } from './dialogs';
import { confirmWrite } from './confirmWrite';
import { UrlBar } from './UrlBar';
import type { WriteConfirmSpec } from '../engine/writeConfirm';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let root: Root | null = null;
let host: HTMLDivElement | null = null;

function mount(children: React.ReactNode) {
  host = document.createElement('div');
  document.body.appendChild(host);
  root = createRoot(host);
  act(() =>
    root!.render(
      <FluentProvider theme={webLightTheme}>
        <AppDialogHost />
        {children}
      </FluentProvider>,
    ),
  );
}

const dialog = () => document.body.querySelector('[role=alertdialog]') as HTMLElement | null;
const button = (text: string) =>
  [...document.body.querySelectorAll('button')].find(
    (b) => b.textContent?.trim() === text,
  ) as HTMLButtonElement;
const flush = () => act(async () => new Promise((r) => setTimeout(r, 0)));

const merge: WriteConfirmSpec = {
  operation: 'Merge Records',
  table: 'account',
  scope: { kind: 'records', count: 2 },
  highImpact: true,
};

beforeEach(() => {
  vi.spyOn(console, 'warn').mockImplementation(() => {});
});
afterEach(() => {
  act(() => root?.unmount());
  host?.remove();
  root = null;
  host = null;
  vi.restoreAllMocks();
});

describe('confirmWrite', () => {
  it('names the environment, operation, table and scope', async () => {
    mount(null);
    let result: Promise<boolean> | null = null;
    act(() => {
      result = confirmWrite(
        merge,
        {
          connectionName: 'Contoso Dev',
          environmentUrl: 'https://contoso.crm.dynamics.com/',
          environment: 'Dev',
        },
        { productionSafeMode: true, verb: 'Send' },
      );
    });
    await flush();
    const text = dialog()?.textContent ?? '';
    expect(text).toContain('Send Merge Records?');
    expect(text).toContain('Contoso Dev');
    expect(text).toContain('Dev');
    expect(text).toContain('contoso.crm.dynamics.com');
    expect(text).toContain('account');
    expect(text).toContain('2 records');
    // Outside Production, a small write isn't typed.
    expect(dialog()?.querySelector('input')).toBeNull();
    act(() => button('Send').click());
    await expect(result!).resolves.toBe(true);
  });

  it('on Production, requires the table name to be typed', async () => {
    mount(null);
    let result: Promise<boolean> | null = null;
    act(() => {
      result = confirmWrite(
        {
          ...merge,
          highImpact: false,
          operation: 'Update Record',
          scope: { kind: 'records', count: 1 },
        },
        {
          connectionName: 'Contoso Prod',
          environmentUrl: 'https://contoso.crm.dynamics.com',
          environment: 'Production',
        },
        { productionSafeMode: true, verb: 'Send' },
      );
    });
    await flush();
    expect(dialog()?.textContent).toContain('This is a Production environment');
    expect(button('Send').disabled).toBe(true);
    const input = dialog()!.querySelector('input') as HTMLInputElement;
    act(() => {
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!.call(
        input,
        'Account',
      );
      input.dispatchEvent(new Event('input', { bubbles: true }));
    });
    expect(button('Send').disabled).toBe(false);
    act(() => button('Send').click());
    await expect(result!).resolves.toBe(true);
  });

  it('resolves false on Cancel', async () => {
    mount(null);
    let result: Promise<boolean> | null = null;
    act(() => {
      result = confirmWrite(
        merge,
        { connectionName: '', environmentUrl: '', environment: '' },
        { productionSafeMode: true, verb: 'Send' },
      );
    });
    await flush();
    act(() => button('Cancel').click());
    await expect(result!).resolves.toBe(false);
  });
});

describe('UrlBar Execute', () => {
  const bar = (onExecute: () => Promise<void>, writeConfirm: WriteConfirmSpec | null) => (
    <UrlBar
      method="POST"
      url="/api/data/v9.2/Merge"
      executeVerb="Send"
      onExecute={onExecute}
      loading={false}
      writeConfirm={writeConfirm}
    />
  );

  it('asks before a high-impact write, and Cancel sends nothing', async () => {
    const onExecute = vi.fn(async () => {});
    mount(bar(onExecute, merge));
    act(() => button('Send').click());
    await flush();
    expect(dialog()).not.toBeNull();
    expect(onExecute).not.toHaveBeenCalled();
    act(() => button('Cancel').click());
    await flush();
    expect(onExecute).not.toHaveBeenCalled();
  });

  it('sends after the confirmation', async () => {
    const onExecute = vi.fn(async () => {});
    mount(bar(onExecute, merge));
    act(() => button('Send').click());
    await flush();
    const confirm = [...dialog()!.querySelectorAll('button')].find(
      (b) => b.textContent?.trim() === 'Send',
    ) as HTMLButtonElement;
    act(() => confirm.click());
    await flush();
    expect(onExecute).toHaveBeenCalledTimes(1);
  });

  it("doesn't ask for a request without a write spec", async () => {
    const onExecute = vi.fn(async () => {});
    mount(bar(onExecute, null));
    act(() => button('Send').click());
    await flush();
    expect(dialog()).toBeNull();
    expect(onExecute).toHaveBeenCalledTimes(1);
  });
});
