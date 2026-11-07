import { lazy, Suspense } from 'react';
import { Spinner } from '@fluentui/react-components';
import type { CodeViewProps } from './CodeView';

// The Code tab, loaded the first time it opens (#236). It brings the code
// generators (every format, the OpenAPI document, the Power Automate pane),
// which most sessions never need at startup. Retrieve Multiple, the mode DRS
// opens on, uses this; the other modes are code-split already.
const LoadedCodeView = lazy(() => import('./CodeView').then((m) => ({ default: m.CodeView })));

export function CodeView(props: CodeViewProps) {
  return (
    <Suspense fallback={<Spinner size="small" label="Loading code…" />}>
      <LoadedCodeView {...props} />
    </Suspense>
  );
}
