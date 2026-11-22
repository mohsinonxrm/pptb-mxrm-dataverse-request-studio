// The Code tab lists the SDK formats (#232) and shows their snippets.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { CodeView } from './CodeView';
import type { CodegenInputs } from '../engine/codeGenerators';

vi.mock('../env/environment', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../env/environment')>()),
  ENV: { name: 'Test', host: 'contoso.crm.dynamics.com', apiBase: '/api/data/v9.2', user: '' },
}));
// Monaco doesn't run in jsdom: show the editor's value as text.
vi.mock('../monaco/MonacoEditor', () => ({
  default: ({ value }: { value?: string }) => <pre data-testid="code">{value}</pre>,
}));

const inputs: CodegenInputs = {
  method: 'GET',
  built: {
    relativeUrl: "/api/data/v9.2/accounts?$select=name&$filter=name eq 'Contoso'",
    relativeNoBase: "/accounts?$select=name&$filter=name eq 'Contoso'",
    bytes: 0,
    queryParts: [
      { key: '$select', value: 'name' },
      { key: '$filter', value: "name eq 'Contoso'" },
    ],
    entitySet: 'accounts',
    entityLogical: 'account',
  },
  headers: {},
};

let hostEl: HTMLDivElement | null = null;
let root: Root | null = null;
beforeEach(() => {
  hostEl = document.createElement('div');
  document.body.appendChild(hostEl);
  root = createRoot(hostEl);
});
afterEach(() => {
  act(() => root?.unmount());
  hostEl?.remove();
  hostEl = null;
});

describe('CodeView — SDK formats', () => {
  it.each([
    ['C# SDK', 'query.Criteria.AddCondition("name", ConditionOperator.Equal, "Contoso");'],
    ['PowerShell (Xrm)', '<condition attribute="name" operator="eq" value="Contoso" />'],
  ])('the %s tab shows its snippet', (label, expected) => {
    act(() =>
      root!.render(
        <FluentProvider theme={webLightTheme}>
          <CodeView themeMode="light" inputs={inputs} />
        </FluentProvider>,
      ),
    );
    expect(hostEl!.textContent).toContain('13 formats');
    const tab = [...hostEl!.querySelectorAll<HTMLButtonElement>('button[role="tab"]')].find(
      // Fluent renders the label twice: a hidden bold copy reserves its width.
      (b) => b.textContent === label.repeat(2),
    );
    expect(tab).toBeDefined();
    act(() => tab!.click());
    expect(hostEl!.querySelector('[data-testid="code"]')?.textContent).toContain(expected);
  });
});
