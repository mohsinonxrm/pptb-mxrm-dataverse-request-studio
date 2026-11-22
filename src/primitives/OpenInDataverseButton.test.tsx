// "Open in Dataverse" on one record (#293, #419): shown only when there's a
// form to open (#422).
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';

const host = vi.hoisted(() => ({ open: vi.fn(async (_url: string) => {}) }));
vi.mock('../host/openInDataverse', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../host/openInDataverse')>()),
  openInBrowser: host.open,
}));
vi.mock('../env/environment', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../env/environment')>()),
  ENV: { name: 'Test', host: 'org.crm.dynamics.com', apiBase: '/api/data/v9.2', user: '' },
}));

import { OpenInDataverseButton, type OpenInDataverseButtonProps } from './OpenInDataverseButton';
import { __clearLiveTables, __registerLiveTable } from '../metadata/metadata';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

const ID = '8038a91f-99c9-f111-aaaf-000d3a5b6760';
let root: Root | null = null;
let el: HTMLDivElement | null = null;

const render = (props: OpenInDataverseButtonProps) =>
  act(() =>
    root!.render(
      <FluentProvider theme={webLightTheme}>
        <OpenInDataverseButton {...props} />
      </FluentProvider>,
    ),
  );
const button = () => el!.querySelector('button');

beforeEach(() => {
  host.open.mockClear();
  el = document.createElement('div');
  document.body.appendChild(el);
  root = createRoot(el);
});
afterEach(() => {
  act(() => root?.unmount());
  el?.remove();
  __clearLiveTables();
});

describe('OpenInDataverseButton', () => {
  it("opens the record's form in the connection's browser", () => {
    render({ table: 'account', id: `{${ID.toUpperCase()}}` });
    expect(button()!.textContent).toBe('Open in Dataverse');
    act(() => button()!.click());
    expect(host.open).toHaveBeenCalledWith(
      `https://org.crm.dynamics.com/main.aspx?pagetype=entityrecord&etn=account&id=${ID}`,
    );
  });

  it('renders nothing without a GUID or a table', () => {
    render({ table: 'account', id: 'nope' });
    expect(button()).toBeNull();
    render({ table: undefined, id: ID });
    expect(button()).toBeNull();
  });

  it('renders nothing for an intersect table (#422)', () => {
    __registerLiveTable({
      logicalName: 'accountleads',
      entitySetName: 'accountleadscollection',
      displayName: 'Account Leads',
      primaryKey: 'accountleadid',
      primaryName: 'name',
      columns: [],
      navigationProperties: [],
      isIntersect: true,
    });
    render({ table: 'accountleads', id: ID });
    expect(button()).toBeNull();
  });

  it('opens an activity as its own type (#422)', () => {
    render({ table: 'activitypointer', id: ID });
    expect(button()).toBeNull();
    render({ table: 'activitypointer', id: ID, record: { activitytypecode: 'task' } });
    act(() => button()!.click());
    expect(host.open).toHaveBeenCalledWith(expect.stringContaining('etn=task&'));
  });
});
