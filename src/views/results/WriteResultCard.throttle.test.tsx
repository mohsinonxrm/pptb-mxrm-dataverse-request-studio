// A write Dataverse throttled (#269) says which limit, and that it didn't run.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { WriteResultCard } from './WriteResultCard';
import { mapError } from '../../engine/dataverseExecutor';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let root: Root | null = null;
let host: HTMLDivElement | null = null;
beforeEach(() => {
  host = document.createElement('div');
  document.body.appendChild(host);
  root = createRoot(host);
});
afterEach(() => {
  act(() => root?.unmount());
  host?.remove();
});

describe('WriteResultCard throttled (#269)', () => {
  it('names the limit and says the write did not run', () => {
    const result = mapError(
      new Error(
        'Dataverse update failed: 0x80072326: Number of concurrent requests exceeded the limit of 52.',
      ),
      4,
    );
    act(() =>
      root!.render(
        <FluentProvider theme={webLightTheme}>
          <WriteResultCard result={result} ctx={{ operation: 'update', table: 'account' }} />
        </FluentProvider>,
      ),
    );
    expect(host!.textContent).toContain('Throttled: concurrent requests');
    expect(host!.textContent).toContain("Dataverse didn't run this update.");
    expect(host!.textContent).toContain('Too many requests were running at once');
  });

  it('names the operation readably', () => {
    const result = mapError(new Error('0x80072322: Number of requests exceeded the limit.'), 4);
    act(() =>
      root!.render(
        <FluentProvider theme={webLightTheme}>
          <WriteResultCard result={result} ctx={{ operation: 'action', table: '' }} />
        </FluentProvider>,
      ),
    );
    expect(host!.textContent).toContain("Dataverse didn't run this action call.");
  });
});
