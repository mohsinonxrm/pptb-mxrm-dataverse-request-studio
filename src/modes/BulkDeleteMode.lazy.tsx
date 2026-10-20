import { lazy, Suspense } from 'react';
import { Spinner } from '@fluentui/react-components';
import type { ThemeMode } from '../theme/theme';

// Bulk Delete Job loads on first use, like the other less-used modes (#22).
// When this lands next to src/modes/lazyModes.tsx it can move into that
// registry.
const Inner = lazy(() => import('./BulkDeleteMode').then((m) => ({ default: m.BulkDeleteMode })));

export function BulkDeleteMode(props: { themeMode: ThemeMode }) {
  return (
    <Suspense
      fallback={
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
          <Spinner size="medium" label="Loading…" />
        </div>
      }
    >
      <Inner {...props} />
    </Suspense>
  );
}
