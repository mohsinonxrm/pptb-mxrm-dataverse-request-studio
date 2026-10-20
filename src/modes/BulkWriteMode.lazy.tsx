import { lazy, Suspense } from 'react';
import { Spinner } from '@fluentui/react-components';
import type { ThemeMode } from '../theme/theme';

// Bulk Write loads on first use, like the other less-used modes (#22): its
// code (CSV/JSON import, row grid) stays off the startup path. When this
// lands next to src/modes/lazyModes.tsx it can move into that registry.
const Inner = lazy(() => import('./BulkWriteMode').then((m) => ({ default: m.BulkWriteMode })));

export function BulkWriteMode(props: { themeMode: ThemeMode }) {
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
