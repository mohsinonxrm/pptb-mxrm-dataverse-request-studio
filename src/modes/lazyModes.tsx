import { lazy, Suspense, type ComponentType } from 'react';
import { Spinner } from '@fluentui/react-components';

// Less-used modes are code-split (#22): each loads the first time it is
// picked, so their code (and editors only they use) stays off the startup
// path. Chunks are served locally by PPTB, so the spinner shows only for a
// moment on first open; later switches render immediately. The everyday
// read/write/execute modes stay in the main bundle.

function lazyMode<P extends object>(
  name: string,
  load: () => Promise<ComponentType<P>>,
): ComponentType<P> {
  // Cast: React.lazy's ref-aware prop typing doesn't line up with a generic P.
  const Lazy = lazy(async () => ({ default: await load() })) as unknown as ComponentType<P>;
  function LazyMode(props: P) {
    return (
      <Suspense
        fallback={
          <div
            style={{
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              minHeight: 0,
            }}
          >
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

export const RetrieveNextLinkMode = lazyMode('RetrieveNextLinkMode', () =>
  import('./RetrieveNextLinkMode').then((m) => m.RetrieveNextLinkMode),
);
export const PredefinedQueryMode = lazyMode('PredefinedQueryMode', () =>
  import('./PredefinedQueryMode').then((m) => m.PredefinedQueryMode),
);
export const MergeMode = lazyMode('MergeMode', () =>
  import('./MergeMode').then((m) => m.MergeMode),
);
export const AssociateMode = lazyMode('AssociateMode', () =>
  import('./AssociateMode').then((m) => m.AssociateMode),
);
export const DisassociateMode = lazyMode('DisassociateMode', () =>
  import('./DisassociateMode').then((m) => m.DisassociateMode),
);
export const ExecuteWorkflowMode = lazyMode('ExecuteWorkflowMode', () =>
  import('./ExecuteWorkflowMode').then((m) => m.ExecuteWorkflowMode),
);
export const ManageFileMode = lazyMode('ManageFileMode', () =>
  import('./ManageFileMode').then((m) => m.ManageFileMode),
);
export const ManageImageMode = lazyMode('ManageImageMode', () =>
  import('./ManageImageMode').then((m) => m.ManageImageMode),
);
