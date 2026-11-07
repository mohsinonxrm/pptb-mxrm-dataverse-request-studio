// useLatestCallback: one identity across renders, the latest closure when called.
import { act } from 'react';
import { createRoot } from 'react-dom/client';
import { describe, expect, it } from 'vitest';
import { useLatestCallback } from './useLatestCallback';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

describe('useLatestCallback', () => {
  it('keeps its identity and calls the latest closure', () => {
    const seen: Array<(x: number) => string> = [];
    function Probe({ label }: { label: string }) {
      seen.push(useLatestCallback((x: number) => `${label}:${x}`));
      return null;
    }
    const root = createRoot(document.createElement('div'));
    act(() => root.render(<Probe label="first" />));
    act(() => root.render(<Probe label="second" />));
    expect(seen).toHaveLength(2);
    expect(seen[1]).toBe(seen[0]);
    // The function kept from the first render sees the second render's data.
    expect(seen[0](1)).toBe('second:1');
    act(() => root.unmount());
  });
});
