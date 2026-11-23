// A request DRS refused before sending says so, with the reason, instead of
// a generic "Failed · 400" that reads as if Dataverse answered (#225).
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { WriteResultCard } from './WriteResultCard';
import { mapError } from '../../engine/dataverseExecutor';
import { BatchPathError } from '../../host/batchUrl';

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

function render(result: ReturnType<typeof mapError>) {
  act(() =>
    root!.render(
      <FluentProvider theme={webLightTheme}>
        <WriteResultCard
          result={result}
          ctx={{ operation: 'upsert', table: 'msfp_unsubscribedrecipient' }}
        />
      </FluentProvider>,
    ),
  );
}

describe('WriteResultCard: not sent', () => {
  it('shows the reason a path was refused before sending', () => {
    render(mapError(new BatchPathError('Power Platform ToolBox refuses this URL.'), 0));
    expect(host!.textContent).toContain('Not sent');
    expect(host!.textContent).toContain('Power Platform ToolBox refuses this URL.');
    expect(host!.textContent).not.toContain('See the JSON tab for details.');
  });

  it('capitalizes a reason that starts lower-case', () => {
    render(
      mapError({ statusCode: 400, message: 'Not sent: the alternate key "k" isn\'t loaded.' }, 0),
    );
    expect(host!.textContent).toContain('The alternate key "k" isn\'t loaded.');
  });

  it('leaves a Dataverse 400 on the generic framing', () => {
    render(mapError({ statusCode: 400, message: 'Bad request body.' }, 12));
    expect(host!.textContent).toContain('Failed · 400');
  });
});
