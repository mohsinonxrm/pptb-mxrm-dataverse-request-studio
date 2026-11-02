// useUrlSuggest on a Fluent Input / Textarea: typing opens suggestions, keys
// move and insert, Esc closes the list without closing a dialog (#268).
import { act, useState } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest';
import {
  Dialog,
  DialogBody,
  DialogSurface,
  FluentProvider,
  Input,
  Textarea,
} from '@fluentui/react-components';
import { registerFixtureTables } from '../test/fixtures';
import type { UrlMode } from '../engine/urlCompletion';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

vi.mock('../host/useScopedEntities', () => {
  const list = [
    { logicalName: 'account', entitySetName: 'accounts', displayName: 'Account' },
    { logicalName: 'contact', entitySetName: 'contacts', displayName: 'Contact' },
  ];
  return { useScopedEntities: () => ({ entities: list, allEntities: list }) };
});
vi.mock('../host/csdlProvider', () => ({
  actions: {
    loadAll: async () => [
      {
        name: 'WhoAmI',
        kind: 'Function',
        namespace: 'Microsoft.Dynamics.CRM',
        binding: { kind: 'unbound' },
        parameters: [],
      },
    ],
  },
}));
vi.mock('../state/metadataScope', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../state/metadataScope')>()),
  useAdvancedFindOnly: () => false,
}));

import { useUrlSuggest } from './useUrlSuggest';

beforeAll(() => {
  registerFixtureTables();
  // jsdom doesn't lay out
  Element.prototype.scrollIntoView ??= () => {};
});

let root: Root;
let host: HTMLDivElement;
const parentKeys: string[] = [];
let dialogOpen = true;

interface HarnessProps {
  multiline?: boolean;
  mode?: UrlMode;
  entitySet?: string;
  initial?: string;
  inDialog?: boolean;
}

function Harness({ multiline, mode = 'url', entitySet, initial = '', inDialog }: HarnessProps) {
  const [value, setValue] = useState(initial);
  const [open, setOpen] = useState(true);
  dialogOpen = open;
  const s = useUrlSuggest({ value, onChange: setValue, mode, entitySet, multiline });
  const field = (
    <div onKeyDown={(e) => parentKeys.push(e.key)}>
      {multiline ? (
        <Textarea value={value} {...s.fieldProps} />
      ) : (
        <Input value={value} {...s.fieldProps} />
      )}
      {s.list}
      <button>elsewhere</button>
    </div>
  );
  if (!inDialog) return field;
  return (
    <Dialog open={open} onOpenChange={(_, d) => setOpen(d.open)}>
      <DialogSurface>
        <DialogBody>{field}</DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

const q = (sel: string) => document.querySelector(sel);
const field = () =>
  q('input:not([type=button]), textarea') as HTMLInputElement | HTMLTextAreaElement;
const options = () =>
  [...document.querySelectorAll('[role=option]')].map((o) => o.getAttribute('aria-label'));
const listOpen = () => !!q('[role=listbox]');
const flush = () =>
  act(async () => {
    await new Promise((r) => setTimeout(r, 0));
  });
async function type(text: string, inputType = 'insertText') {
  const el = field();
  await act(async () => {
    const proto = el instanceof HTMLTextAreaElement ? HTMLTextAreaElement : HTMLInputElement;
    const at = el.selectionStart ?? el.value.length;
    const next = el.value.slice(0, at) + text + el.value.slice(at);
    Object.getOwnPropertyDescriptor(proto.prototype, 'value')!.set!.call(el, next);
    el.setSelectionRange(at + text.length, at + text.length);
    el.dispatchEvent(new InputEvent('input', { bubbles: true, inputType }));
  });
}
const press = (key: string, init: KeyboardEventInit = {}) =>
  act(() => {
    field().dispatchEvent(
      new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true, ...init }),
    );
  });

beforeEach(() => {
  parentKeys.length = 0;
  host = document.createElement('div');
  document.body.appendChild(host);
  root = createRoot(host);
});
afterEach(() => {
  act(() => root.unmount());
  host.remove();
});

async function mount(props: HarnessProps = {}) {
  await act(async () =>
    root.render(
      <FluentProvider>
        <Harness {...props} />
      </FluentProvider>,
    ),
  );
  await act(async () => field().focus());
  await flush();
}

describe('useUrlSuggest', () => {
  it('opens on an edit, lists unbound functions, and inserts the active one', async () => {
    await mount();
    // Focus alone doesn't open it
    expect(listOpen()).toBe(false);
    await type('Who');
    expect(options()).toEqual(['WhoAmI, Function']);
    expect(field().getAttribute('aria-expanded')).toBe('true');
    expect(field().getAttribute('aria-activedescendant')).toBe(q('[role=option]')!.id);
    expect(q('[role=status]')!.textContent).toBe('1 suggestion');
    await press('Enter');
    expect(field().value).toBe('WhoAmI()');
    expect(listOpen()).toBe(false);
    expect(q('[role=status]')!.textContent).toBe('');
  });

  it('keeps going after an option name, and moves with the arrow keys', async () => {
    await mount();
    await type('accounts?$sel');
    expect(options()).toEqual(['$select']);
    await press('Tab');
    expect(field().value).toBe('accounts?$select=');
    expect(field().selectionStart).toBe('accounts?$select='.length);
    expect(options()[0]).toBe('accountid, Account (Uniqueidentifier)');
    await press('ArrowDown');
    await press('Enter');
    expect(field().value).toBe('accounts?$select=accountnumber');
  });

  it('starts a narrowed list at its best match', async () => {
    await mount();
    await type('accounts?$select=');
    await press('Escape');
    await type('a');
    await press('ArrowDown');
    await type('ccountn');
    await press('Enter');
    expect(field().value).toBe('accounts?$select=accountnumber');
  });

  it('moves the caret on an insert that leaves the text as it was', async () => {
    // Mid-word in `$sel|ect=`: inserting `$select=` changes nothing
    await mount({ initial: 'accounts?$select=' });
    await act(async () => field().setSelectionRange(13, 13));
    await press(' ', { ctrlKey: true });
    expect(options()).toEqual(['$select']);
    await press('Tab');
    expect(field().value).toBe('accounts?$select=');
    expect(field().selectionStart).toBe(17);
    // The next edits land after it, and the columns follow
    await type('na');
    expect(field().value).toBe('accounts?$select=na');
    expect(options()).toEqual(['name, Name (String)']);
  });

  it('inserts on click, and closes on blur', async () => {
    await mount();
    await type('cont');
    await act(async () => (q('[role=option]') as HTMLElement).click());
    expect(field().value).toBe('contacts');
    await type('?$t');
    expect(listOpen()).toBe(true);
    await act(async () => (q('button') as HTMLButtonElement).focus());
    expect(listOpen()).toBe(false);
  });

  it('doesn’t open on a paste, or insert on Ctrl+Enter', async () => {
    await mount();
    await type('accounts?$select=na', 'insertFromPaste');
    expect(listOpen()).toBe(false);
    await type('m');
    expect(listOpen()).toBe(true);
    await press('Enter', { ctrlKey: true });
    expect(field().value).toBe('accounts?$select=nam');
    expect(parentKeys).toContain('Enter');
  });

  it('completes query options against the table in query mode', async () => {
    await mount({ mode: 'query', entitySet: 'contacts' });
    await type('$select=full');
    expect(options()).toEqual(['fullname, Full Name (String)']);
  });

  it('works in a Textarea, without the combobox role a textarea can’t take', async () => {
    await mount({ multiline: true });
    await type('contacts?$select=full');
    expect(options()).toEqual(['fullname, Full Name (String)']);
    expect(field().getAttribute('role')).toBeNull();
    expect(field().getAttribute('aria-controls')).toBe(q('[role=listbox]')!.id);
    await press('Enter');
    expect(field().value).toBe('contacts?$select=fullname');
  });

  it('closes the list on Esc in a Fluent Dialog, and the dialog on the next Esc', async () => {
    await mount({ inDialog: true });
    await type('acc');
    expect(listOpen()).toBe(true);
    await press('Escape');
    expect(listOpen()).toBe(false);
    expect(dialogOpen).toBe(true);
    await press('Escape');
    expect(dialogOpen).toBe(false);
  });
});
