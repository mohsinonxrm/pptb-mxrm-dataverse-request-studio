// #387: the status pill is a Fluent v9 filled Badge, with the check icon on
// success, as the v2.2 design maps it.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { StatusPill } from './StatusPill';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let root: Root;
let host: HTMLDivElement;
beforeEach(() => {
  host = document.createElement('div');
  root = createRoot(host);
});
afterEach(() => act(() => root.unmount()));

const render = (el: React.ReactElement) =>
  act(() => root.render(<FluentProvider theme={webLightTheme}>{el}</FluentProvider>));
const badge = () => host.querySelector<HTMLElement>('.fui-Badge')!;

describe('StatusPill', () => {
  it('is a Fluent Badge with the reason phrase, the time and a check on success', () => {
    render(<StatusPill status="success" code={201} ms={42} />);
    expect(badge()).not.toBeNull();
    expect(badge().textContent).toBe('201 Created · 42 ms');
    expect(badge().querySelector('svg')).not.toBeNull();
  });

  it('shows the code alone, without the check, for other statuses', () => {
    render(<StatusPill status="danger" code={404} />);
    expect(badge().textContent).toBe('404');
    expect(badge().querySelector('svg')).toBeNull();
  });
});
