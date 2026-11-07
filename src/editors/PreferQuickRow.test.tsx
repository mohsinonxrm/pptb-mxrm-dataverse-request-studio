// #299: the Prefer toggles under the URL bar for write types.
import { act } from 'react';
import { createRoot } from 'react-dom/client';
import { describe, expect, it, vi } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { PreferQuickRow, preferEffectLine } from './PreferQuickRow';
import { emptyPrefer } from './PreferEditor';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

describe('preferEffectLine', () => {
  it('says what each choice returns', () => {
    const off = emptyPrefer();
    const on = { ...off, returnRepresentation: true, formattedValues: true };
    expect(preferEffectLine('create', off)).toContain('204 No Content');
    expect(preferEffectLine('create', off)).toContain('OData-EntityId');
    expect(preferEffectLine('create', on)).toBe(
      'Returns 201 Created with the new record, with formatted values.',
    );
    expect(preferEffectLine('update', { ...on, formattedValues: false })).toBe(
      'Returns 200 OK with the updated record.',
    );
    expect(preferEffectLine('upsert', on)).toContain('created or updated record');
  });
});

describe('PreferQuickRow', () => {
  it('edits the same Prefer spec', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const root = createRoot(host);
    const setSpec = vi.fn();
    act(() =>
      root.render(
        <FluentProvider theme={webLightTheme}>
          <PreferQuickRow operation="create" spec={emptyPrefer()} setSpec={setSpec} />
        </FluentProvider>,
      ),
    );
    const [rep, formatted] = [...host.querySelectorAll<HTMLInputElement>('input[type=checkbox]')];
    expect(formatted.disabled).toBe(true); // nothing comes back to format
    act(() => rep.click());
    expect(setSpec).toHaveBeenCalledWith(expect.objectContaining({ returnRepresentation: true }));
    act(() => root.unmount());
    host.remove();
  });
});
