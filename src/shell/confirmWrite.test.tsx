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
import type { HostSession } from '../host/pptbBridge';

// The connection UrlBar reads when Execute is pressed. Tests change it to
// stand for a Production connection, or a switch while the dialog is open.
const session = vi.hoisted(() => ({
  current: {
    embedded: true,
    connected: true,
    theme: 'light',
    environmentUrl: 'https://contoso.crm.dynamics.com',
    environment: 'Dev',
    connectionName: 'Contoso Sandbox',
  } as Partial<HostSession>,
}));
vi.mock('../host/pptbBridge', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../host/pptbBridge')>()),
  getSession: () => session.current,
}));

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
const dialogButton = (text: string) =>
  [...dialog()!.querySelectorAll('button')].find(
    (b) => b.textContent?.trim() === text,
  ) as HTMLButtonElement;
const flush = () => act(async () => new Promise((r) => setTimeout(r, 0)));
const type = (input: HTMLInputElement, value: string) =>
  act(() => {
    Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!.call(input, value);
    input.dispatchEvent(new Event('input', { bubbles: true }));
  });
const press = (el: Element, key: string, init: KeyboardEventInit = {}) =>
  act(() => {
    el.dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, ...init }));
  });

const merge: WriteConfirmSpec = {
  operation: 'Merge Records',
  table: 'account',
  scope: { kind: 'records', count: 2 },
  highImpact: true,
};
const update: WriteConfirmSpec = {
  operation: 'Update Record',
  table: 'account',
  scope: { kind: 'records', count: 1 },
  highImpact: false,
};
const dev = {
  connectionName: 'Contoso Sandbox',
  environmentUrl: 'https://contoso.crm.dynamics.com/',
  environment: 'Dev',
};
const prod = {
  connectionName: 'Contoso Live',
  environmentUrl: 'https://contoso.crm.dynamics.com',
  environment: 'Production',
};

beforeEach(() => {
  vi.spyOn(console, 'warn').mockImplementation(() => {});
  session.current = {
    ...session.current,
    ...dev,
    environmentUrl: 'https://contoso.crm.dynamics.com',
  };
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
      result = confirmWrite(merge, dev, { productionSafeMode: true, verb: 'Send' });
    });
    await flush();
    const text = dialog()?.textContent ?? '';
    expect(text).toContain('Confirm: Merge Records');
    expect(text).toContain('Contoso Sandbox');
    // The environment type, as its own badge.
    expect([...dialog()!.querySelectorAll('.fui-Badge')].map((b) => b.textContent)).toContain(
      'Dev',
    );
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
      result = confirmWrite(update, prod, { productionSafeMode: true, verb: 'Send' });
    });
    await flush();
    expect(dialog()?.textContent).toContain('This is a Production environment');
    expect(button('Send').disabled).toBe(true);
    type(dialog()!.querySelector('input') as HTMLInputElement, 'Account');
    expect(button('Send').disabled).toBe(false);
    act(() => button('Send').click());
    await expect(result!).resolves.toBe(true);
  });

  it('requires typing for more than 100 records outside Production too', async () => {
    mount(null);
    act(() => {
      void confirmWrite(
        { ...merge, operation: 'Bulk Write', scope: { kind: 'records', count: 101 } },
        dev,
        { productionSafeMode: true, verb: 'Send' },
      );
    });
    await flush();
    expect(dialog()?.querySelector('input')).not.toBeNull();
    expect(button('Send').disabled).toBe(true);
    act(() => button('Cancel').click());
  });

  it('Enter confirms only once the word matches, and not while an IME is composing', async () => {
    mount(null);
    let settled: boolean | null = null;
    act(() => {
      void confirmWrite(update, prod, { productionSafeMode: true, verb: 'Send' }).then(
        (ok) => (settled = ok),
      );
    });
    await flush();
    const input = dialog()!.querySelector('input') as HTMLInputElement;
    type(input, 'contact');
    press(input, 'Enter');
    await flush();
    expect(settled).toBeNull();
    expect(dialog()).not.toBeNull();
    type(input, 'account');
    press(input, 'Enter', { isComposing: true });
    await flush();
    expect(settled).toBeNull();
    press(input, 'Enter');
    await flush();
    expect(settled).toBe(true);
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
    act(() => dialogButton('Send').click());
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

  it("doesn't ask for a small write on a Dev connection", async () => {
    const onExecute = vi.fn(async () => {});
    mount(bar(onExecute, update));
    act(() => button('Send').click());
    await flush();
    expect(dialog()).toBeNull();
    expect(onExecute).toHaveBeenCalledTimes(1);
  });

  it('asks for the same write on a Production connection, with the word typed', async () => {
    session.current = { ...session.current, ...prod };
    const onExecute = vi.fn(async () => {});
    mount(bar(onExecute, update));
    act(() => button('Send').click());
    await flush();
    expect(dialog()?.textContent).toContain('Contoso Live');
    expect(dialogButton('Send').disabled).toBe(true);
    type(dialog()!.querySelector('input') as HTMLInputElement, 'account');
    act(() => dialogButton('Send').click());
    await flush();
    expect(onExecute).toHaveBeenCalledTimes(1);
  });

  it("doesn't send when the connection changes while the confirmation is open", async () => {
    const onExecute = vi.fn(async () => {});
    mount(bar(onExecute, merge));
    act(() => button('Send').click());
    await flush();
    expect(dialog()?.textContent).toContain('Contoso Sandbox');
    session.current = {
      ...session.current,
      ...prod,
      environmentUrl: 'https://fabrikam.crm.dynamics.com',
    };
    act(() => dialogButton('Send').click());
    await flush();
    expect(onExecute).not.toHaveBeenCalled();
    expect(document.body.textContent).toContain('The connection changed');
  });

  it("doesn't send when Escape closes the confirmation", async () => {
    const onExecute = vi.fn(async () => {});
    mount(bar(onExecute, merge));
    act(() => button('Send').click());
    await flush();
    press(dialog()!, 'Escape');
    await flush();
    expect(dialog()).toBeNull();
    expect(onExecute).not.toHaveBeenCalled();
  });
});
