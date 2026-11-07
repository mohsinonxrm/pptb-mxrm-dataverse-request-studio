// #290: one summary line at the top of every results view.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { formatBytes, ResultsSummaryBar, statusKindOf } from './ResultsSummaryBar';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let root: Root;
let host: HTMLDivElement;
beforeEach(() => {
  host = document.createElement('div');
  root = createRoot(host);
});
afterEach(() => act(() => root.unmount()));

const text = (el: React.ReactElement) => {
  act(() => root.render(<FluentProvider theme={webLightTheme}>{el}</FluentProvider>));
  return host.querySelector('[data-testid=results-summary]')!.textContent!.replace(/\s+/g, ' ');
};

/** The pieces appear, in this order. */
const inOrder = (got: string, pieces: string[]) => {
  let at = -1;
  for (const p of pieces) {
    const i = got.indexOf(p, at + 1);
    expect([p, i > at]).toEqual([p, true]);
    at = i;
  }
};

describe('ResultsSummaryBar', () => {
  it('reads status, time, size, rows, total and more pages in order', () => {
    inOrder(
      text(<ResultsSummaryBar code={200} ms={238} bytes={4300} rows={12} total={1342} more />),
      [
        '200 OK',
        '238 ms',
        '4.2 KB',
        '12 rows on this page',
        `total ${(1342).toLocaleString()}`,
        'more pages',
      ],
    );
  });

  it('leaves out what a response doesn’t have', () => {
    const got = text(<ResultsSummaryBar code={204} ms={50} bytes={0} />);
    inOrder(got, ['204 No Content', '50 ms', '0.0 KB']);
    expect(got).not.toMatch(/row|total|more/);
    expect(text(<ResultsSummaryBar code={200} rows={1} />)).toMatch(/1 row$/);
  });

  it('colours by status and formats sizes', () => {
    expect([statusKindOf(200), statusKindOf(304), statusKindOf(404), statusKindOf(0)]).toEqual([
      'success',
      'warning',
      'danger',
      'danger',
    ]);
    expect(formatBytes(2.5 * 1024 * 1024)).toBe('2.5 MB');
  });
});
