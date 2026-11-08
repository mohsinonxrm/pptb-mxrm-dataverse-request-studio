// #248: Custom APIs / Custom Actions list your org's and ISVs' operations
// above Microsoft's, each with its publisher; OOB actions stay ungrouped.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import type { CsdlAction } from '../mock/actionsCsdl';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const op = (name: string, extra: Partial<CsdlAction> = {}): CsdlAction => ({
  name,
  kind: 'Action',
  namespace: 'Microsoft.Dynamics.CRM',
  binding: { kind: 'unbound' },
  parameters: [],
  returnType: { kind: 'void', typeName: 'void' } as CsdlAction['returnType'],
  source: 'custom-api',
  ...extra,
});

const CATALOG: CsdlAction[] = [
  op('msdyn_Summarize', { publisher: { name: 'Dynamics 365', microsoft: true } }),
  op('contoso_DoThing', { publisher: { name: 'Contoso', microsoft: false } }),
  op('pm_Analyze', { publisher: { name: 'Microsoft First Party', microsoft: true } }),
  op('fab_Sync', { publisher: { name: 'Fabrikam ISV', microsoft: false } }),
  op('WinOpportunity', { source: 'oob' }),
  op('QualifyLead', { source: 'oob' }),
  // Execute Function lists OOB functions and Custom API functions together.
  op('WhoAmI', { kind: 'Function', source: 'oob' }),
  op('contoso_GetScore', {
    kind: 'Function',
    publisher: { name: 'Contoso', microsoft: false },
  }),
  op('msdyn_GetKpi', {
    kind: 'Function',
    publisher: { name: 'Dynamics 365', microsoft: true },
  }),
];

vi.mock('../host/csdlProvider', () => ({
  actions: { loadAll: async () => CATALOG },
  isCsdlSizeLimitError: () => false,
  CSDL_SIZE_LIMIT_ISSUE_URL: '',
}));

import { ActionPicker } from './ActionPicker';

let root: Root | null = null;
let host: HTMLDivElement | null = null;

afterEach(() => {
  act(() => root?.unmount());
  host?.remove();
  root = null;
  host = null;
});

const oob = (a: CsdlAction) => a.source === 'oob' && a.kind === 'Action';
const customApis = (a: CsdlAction) => a.source === 'custom-api' && a.kind === 'Action';
const functions = (a: CsdlAction) => a.kind === 'Function';

/** Mount with `value` picked (its Unbound scope is derived), then open the Operation list. */
async function openOperations(filter: (a: CsdlAction) => boolean, value: string) {
  host = document.createElement('div');
  document.body.appendChild(host);
  root = createRoot(host);
  await act(async () => {
    root!.render(
      <FluentProvider theme={webLightTheme}>
        <ActionPicker filter={filter} value={value} title="Pick" onChange={() => undefined} />
      </FluentProvider>,
    );
  });
  const inputs = [...host.querySelectorAll<HTMLInputElement>('input[role=combobox]')];
  const operation = inputs[1];
  await act(async () => {
    operation.click();
  });
  return document.body;
}

const groupLabels = (body: HTMLElement) =>
  [...body.querySelectorAll('[role=group]')].map((g) => {
    const id = g.getAttribute('aria-labelledby');
    return (id ? document.getElementById(id)?.textContent : g.textContent)?.trim();
  });
const optionNames = (body: HTMLElement) =>
  [...body.querySelectorAll('[role=option]')].map((o) => o.textContent ?? '');
/** An option's operation name: its text up to the "Action" kind badge. */
const opName = (text: string) => text.match(/^(\w+?)(Action|Function)/)?.[1];

describe('ActionPicker publisher groups (#248)', () => {
  it('lists your org & ISVs above Microsoft, with a publisher badge each', async () => {
    const body = await openOperations(customApis, 'contoso_DoThing');
    expect(groupLabels(body)).toEqual(['Your org & ISVs (2)', 'Microsoft (2)']);
    const names = optionNames(body);
    expect(names.map((n) => opName(n))).toEqual([
      'contoso_DoThing',
      'fab_Sync',
      'msdyn_Summarize',
      'pm_Analyze',
    ]);
    expect(names[0]).toContain('Contoso');
    expect(names[3]).toContain('Microsoft First Party');
    // The selected operation's card names its publisher too.
    expect(host!.textContent).toMatch(/PublisherContoso/);
  });

  it('counts your own operations in the scope list', async () => {
    host = document.createElement('div');
    document.body.appendChild(host);
    root = createRoot(host);
    await act(async () => {
      root!.render(
        <FluentProvider theme={webLightTheme}>
          <ActionPicker filter={customApis} value={null} title="Pick" onChange={() => undefined} />
        </FluentProvider>,
      );
    });
    const scope = host.querySelector<HTMLInputElement>('input[role=combobox]')!;
    await act(async () => {
      scope.click();
    });
    expect(optionNames(document.body).join(' ')).toContain('4 ops, 2 yours');
  });

  it('files OOB functions under Microsoft when the list mixes them with Custom APIs', async () => {
    const body = await openOperations(functions, 'WhoAmI');
    expect(groupLabels(body)).toEqual(['Your org & ISVs (1)', 'Microsoft (2)']);
    expect(optionNames(body).map((n) => opName(n))).toEqual([
      'contoso_GetScore',
      'WhoAmI',
      'msdyn_GetKpi',
    ]);
  });

  it('leaves OOB actions ungrouped', async () => {
    const body = await openOperations(oob, 'WinOpportunity');
    expect(groupLabels(body)).toEqual([]);
    expect(optionNames(body).map((n) => opName(n))).toEqual(['WinOpportunity', 'QualifyLead']);
  });
});
