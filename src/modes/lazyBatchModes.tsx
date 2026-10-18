import { lazy, Suspense, type ComponentType } from 'react';
import { Spinner } from '@fluentui/react-components';
import type { ThemeMode } from '../theme/theme';

// The Batch group's modes load on first use, so their code stays off the
// startup path. Same pattern as the #22 lazy-mode loader; fold these into
// `lazyModes.tsx` once that lands.

type ModeProps = { themeMode: ThemeMode };

function lazyMode(name: string, load: () => Promise<ComponentType<ModeProps>>) {
  const Lazy = lazy(async () => ({ default: await load() }));
  function LazyMode(props: ModeProps) {
    return (
      <Suspense
        fallback={
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
            <Spinner size="medium" label="Loading…" />
          </div>
        }
      >
        <Lazy {...props} />
      </Suspense>
    );
  }
  LazyMode.displayName = name;
  return LazyMode;
}

export const BatchMode = lazyMode('BatchMode', () =>
  import('./BatchMode').then((m) => m.BatchMode),
);
