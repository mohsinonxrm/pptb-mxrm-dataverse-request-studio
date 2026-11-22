// The $apply editor renders before its table's metadata has loaded, then
// again once it has (#273: ESLint's rules-of-hooks found two hooks after an
// early return, which made React throw when the table arrived).
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, describe, expect, it } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { ApplyEditor, emptyApply } from './ApplyEditor';
import { __clearLiveTables } from '../metadata/metadata';
import { registerFixtureTables } from '../test/fixtures';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let root: Root | null = null;
let host: HTMLDivElement | null = null;
afterEach(() => {
  act(() => root?.unmount());
  host?.remove();
  root = null;
  host = null;
});

const render = () =>
  act(() =>
    root!.render(
      <FluentProvider theme={webLightTheme}>
        <ApplyEditor table="account" spec={emptyApply()} setSpec={() => {}} />
      </FluentProvider>,
    ),
  );

describe('ApplyEditor', () => {
  it('renders once the table arrives after a render without it', () => {
    __clearLiveTables();
    host = document.createElement('div');
    document.body.appendChild(host);
    root = createRoot(host);
    render();
    expect(host.textContent).toBe('');
    registerFixtureTables();
    expect(render).not.toThrow();
    expect(host.textContent).toContain('$apply');
  });
});
