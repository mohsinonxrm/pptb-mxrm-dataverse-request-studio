// #243: the sidebar's clauses are a nav of buttons in lists, nested clauses
// in nested lists. No tree roles (it has no tree keyboard model) and no
// aria-expanded (nothing collapses).
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { Filter20Regular, Table20Regular } from '@fluentui/react-icons';
import { Sidebar, type SidebarSection } from './Sidebar';
import { findRequestType } from '../registry/requestTypes';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const SECTIONS: SidebarSection[] = [
  {
    id: 'target',
    label: 'Target',
    items: [{ id: 'target', icon: Table20Regular, label: 'Account', badge: '✓' }],
  },
  {
    id: 'query',
    label: 'Query structure',
    meta: '2 active',
    items: [
      {
        id: 'filter',
        icon: Filter20Regular,
        label: '$filter',
        code: true,
        badge: '2',
        dirty: true,
        children: [
          { id: 'filter:0', icon: Filter20Regular, label: 'name eq x' },
          {
            id: 'filter:1',
            icon: Filter20Regular,
            label: 'or group',
            children: [{ id: 'filter:1:0', icon: Filter20Regular, label: 'city eq y' }],
          },
        ],
      },
      { id: 'top', icon: Filter20Regular, label: '$top', code: true, badge: '50' },
    ],
  },
];

let root: Root | null = null;
let host: HTMLDivElement | null = null;
afterEach(() => {
  act(() => root?.unmount());
  host?.remove();
  root = null;
  host = null;
});

function mount(activeNode: string, onSelect = vi.fn()) {
  host = document.createElement('div');
  document.body.appendChild(host);
  root = createRoot(host);
  act(() =>
    root!.render(
      <FluentProvider theme={webLightTheme}>
        <Sidebar
          type={findRequestType('retrieve-multiple')}
          urlPreview="/accounts"
          sections={SECTIONS}
          activeNode={activeNode}
          onSelect={onSelect}
          recents={[]}
        />
      </FluentProvider>,
    ),
  );
  return { host: host!, onSelect };
}

const button = (el: HTMLElement, text: string) => {
  const b = [...el.querySelectorAll('button')].find((b) => b.textContent!.includes(text));
  if (!b) throw new Error('no button ' + text);
  return b;
};

describe('Sidebar clause list (#243)', () => {
  it('uses a nav of lists and buttons, with no tree roles or expand state', () => {
    const { host } = mount('top');
    const nav = host.querySelector('nav')!;
    expect(nav.getAttribute('aria-label')).toBe('Request clauses');
    expect(host.querySelector('[role=tree], [role=treeitem], [role=group]')).toBeNull();
    expect(host.querySelector('[aria-expanded], [aria-selected]')).toBeNull();
    // Every list item is an <li> directly in a <ul>.
    for (const li of nav.querySelectorAll('li')) expect(li.parentElement!.tagName).toBe('UL');
    for (const ul of nav.querySelectorAll('ul'))
      for (const c of ul.children) expect(c.tagName).toBe('LI');
  });

  it('names each section list by its heading', () => {
    const { host } = mount('top');
    const lists = [...host.querySelectorAll('nav > div > ul')];
    expect(
      lists.map((ul) => document.getElementById(ul.getAttribute('aria-labelledby')!)?.textContent),
    ).toEqual(['Target', 'Query structure']);
  });

  it('nests child clauses in nested lists, one level per depth', () => {
    const { host } = mount('top');
    const depth = (b: HTMLElement) => {
      let n = 0;
      for (let e: HTMLElement | null = b; e; e = e.parentElement) if (e.tagName === 'UL') n++;
      return n;
    };
    expect(depth(button(host, '$filter'))).toBe(1);
    expect(depth(button(host, 'name eq x'))).toBe(2);
    expect(depth(button(host, 'or group'))).toBe(2);
    expect(depth(button(host, 'city eq y'))).toBe(3);
    expect(depth(button(host, '$top'))).toBe(1);
    // The nested list belongs to its parent's item.
    expect(button(host, 'name eq x').closest('ul')!.parentElement).toBe(
      button(host, '$filter').closest('li'),
    );
  });

  it('marks the open clause with aria-current and selects on click', () => {
    const { host, onSelect } = mount('filter:1');
    const current = [...host.querySelectorAll('[aria-current]')];
    expect(current).toHaveLength(1);
    expect(current[0].getAttribute('aria-current')).toBe('true');
    expect(current[0].textContent).toContain('or group');
    act(() => button(host, '$top').click());
    expect(onSelect).toHaveBeenCalledWith('top');
  });

  it('says "modified" and "Set" in a way screen readers read', () => {
    const { host } = mount('top');
    const filter = button(host, '$filter');
    expect(filter.textContent).toContain('(modified)');
    expect(filter.querySelector('[aria-label]')).toBeNull();
    const set = button(host, 'Account').querySelector('[role=img]')!;
    expect(set.getAttribute('aria-label')).toBe('Set');
  });
});
