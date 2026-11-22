import { lazy, Suspense, type ComponentType } from 'react';
import { Spinner } from '@fluentui/react-components';

// Every mode but Retrieve Multiple, the one DRS opens on, is code-split
// (#22, #236): each loads the first time it is picked, so its code (and the
// editors only it uses) stays off the startup path. Chunks are served locally
// by PPTB, so the spinner shows only for a moment on first open; later
// switches render immediately.

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
export const RetrieveSingleMode = lazyMode('RetrieveSingleMode', () =>
  import('./RetrieveSingleMode').then((m) => m.RetrieveSingleMode),
);
export const CreateMode = lazyMode('CreateMode', () =>
  import('./CreateMode').then((m) => m.CreateMode),
);
export const UpdateMode = lazyMode('UpdateMode', () =>
  import('./UpdateMode').then((m) => m.UpdateMode),
);
export const UpsertMode = lazyMode('UpsertMode', () =>
  import('./UpsertMode').then((m) => m.UpsertMode),
);
export const DeleteMode = lazyMode('DeleteMode', () =>
  import('./DeleteMode').then((m) => m.DeleteMode),
);
export const ExecuteActionMode = lazyMode('ExecuteActionMode', () =>
  import('./ExecuteActionMode').then((m) => m.ExecuteActionMode),
);
export const ExecuteFunctionMode = lazyMode('ExecuteFunctionMode', () =>
  import('./ExecuteFunctionMode').then((m) => m.ExecuteFunctionMode),
);
export const ManageAttachmentMode = lazyMode('ManageAttachmentMode', () =>
  import('./ManageAttachmentMode').then((m) => m.ManageAttachmentMode),
);
