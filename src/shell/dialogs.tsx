// In-app replacements for the browser's native confirm / alert boxes (#111).
//
// Inside Power Platform ToolBox a native confirm or alert box opens a Windows
// message box that blocks the DRS renderer and PPTB's own popup windows
// until it is answered. Everything here renders as Fluent UI v9 in the DRS
// DOM instead, so nothing outside the tool is ever blocked.
//
//   <AppDialogHost />  — mounted once at the app root (inside FluentProvider).
//                         Renders the pending confirm dialog and the toaster.
//   useConfirm()       — hook returning `(options) => Promise<boolean>`.
//   confirmDialog()    — the same, callable outside React components.
//   notify()           — non-blocking toast; drop-in for `window.alert`.
//
// The functions talk to the host through a small module-level bridge, so
// callers need no context or provider of their own. Calls made before the
// host mounts are queued and shown once it does.

import { useCallback, useEffect, useState, type ReactNode } from 'react';
import {
  Button,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Toast,
  ToastBody,
  ToastTitle,
  ToastTrigger,
  Toaster,
  makeStyles,
  tokens,
  useId,
  useToastController,
  type ToastIntent,
} from '@fluentui/react-components';
import { Dismiss16Regular } from '@fluentui/react-icons';

export interface ConfirmOptions {
  /** Dialog title. Defaults to "Are you sure?". */
  title?: string;
  /** Body text or content. */
  message?: ReactNode;
  /** Confirm button label. Defaults to "OK", or "Delete" when destructive. */
  confirmLabel?: string;
  /** Cancel button label. Defaults to "Cancel". */
  cancelLabel?: string;
  /** Style the confirm button as a destructive (red) action. */
  destructive?: boolean;
}

export interface NotifyOptions {
  /** Bold first line of the toast (it carries the intent icon). Defaults by intent. */
  title?: string;
  /** Defaults to "warning", which suits the former alert call sites. */
  intent?: ToastIntent;
  /** Auto-dismiss after this many ms. Defaults to 10 s; -1 keeps it open. */
  timeout?: number;
}

interface PendingConfirm {
  id: number;
  options: ConfirmOptions;
  resolve: (ok: boolean) => void;
}

interface PendingNotice {
  message: ReactNode;
  options: NotifyOptions;
}

interface HostBridge {
  ask: (p: PendingConfirm) => void;
  notify: (n: PendingNotice) => void;
}

let bridge: HostBridge | null = null;
let queuedConfirms: PendingConfirm[] = [];
let queuedNotices: PendingNotice[] = [];
let nextId = 1;

/** Ask the user to confirm. Resolves true on confirm, false on cancel / dismiss. */
export function confirmDialog(options: ConfirmOptions | string): Promise<boolean> {
  const opts = typeof options === 'string' ? { message: options } : options;
  return new Promise<boolean>((resolve) => {
    const pending: PendingConfirm = { id: nextId++, options: opts, resolve };
    if (bridge) bridge.ask(pending);
    else queuedConfirms.push(pending);
  });
}

/** Hook form of `confirmDialog`, for components. */
export function useConfirm(): (options: ConfirmOptions | string) => Promise<boolean> {
  return useCallback((options: ConfirmOptions | string) => confirmDialog(options), []);
}

/** Show a non-blocking in-app notice. Drop-in replacement for `window.alert`. */
export function notify(message: ReactNode, options: NotifyOptions = {}): void {
  const notice: PendingNotice = { message, options };
  if (bridge) bridge.notify(notice);
  else queuedNotices.push(notice);
}

const DEFAULT_TITLES: Record<ToastIntent, string> = {
  info: 'Note',
  success: 'Done',
  warning: 'Warning',
  error: 'Error',
};

const useStyles = makeStyles({
  message: {
    whiteSpace: 'pre-wrap',
    overflowWrap: 'anywhere',
  },
  destructive: {
    backgroundColor: tokens.colorStatusDangerBackground3,
    color: tokens.colorNeutralForegroundOnBrand,
    ':hover': {
      backgroundColor: tokens.colorStatusDangerBackground3Hover,
      color: tokens.colorNeutralForegroundOnBrand,
    },
    ':hover:active': {
      backgroundColor: tokens.colorStatusDangerBackground3Pressed,
      color: tokens.colorNeutralForegroundOnBrand,
    },
  },
});

/** Mount once at the app root, inside the FluentProvider. */
export function AppDialogHost() {
  const s = useStyles();
  const toasterId = useId('drs-toaster');
  const { dispatchToast } = useToastController(toasterId);
  const [confirms, setConfirms] = useState<PendingConfirm[]>([]);

  useEffect(() => {
    const showNotice = ({ message, options }: PendingNotice) => {
      const intent = options.intent ?? 'warning';
      dispatchToast(
        <Toast>
          <ToastTitle
            action={
              <ToastTrigger>
                <Button
                  appearance="transparent"
                  size="small"
                  icon={<Dismiss16Regular />}
                  aria-label="Dismiss"
                />
              </ToastTrigger>
            }
          >
            {options.title ?? DEFAULT_TITLES[intent]}
          </ToastTitle>
          <ToastBody className={s.message}>{message}</ToastBody>
        </Toast>,
        {
          intent,
          timeout: options.timeout ?? 10_000,
          pauseOnHover: true,
          pauseOnWindowBlur: true,
        },
      );
    };
    bridge = {
      ask: (p) => setConfirms((q) => [...q, p]),
      notify: showNotice,
    };
    if (queuedConfirms.length > 0) {
      const flushed = queuedConfirms;
      queuedConfirms = [];
      setConfirms((q) => [...q, ...flushed]);
    }
    if (queuedNotices.length > 0) {
      const flushed = queuedNotices;
      queuedNotices = [];
      flushed.forEach(showNotice);
    }
    return () => {
      bridge = null;
    };
  }, [dispatchToast, s.message]);

  const current = confirms[0];
  const settle = (ok: boolean) => {
    if (!current) return;
    current.resolve(ok);
    setConfirms((q) => q.filter((p) => p.id !== current.id));
  };

  const o = current?.options;
  return (
    <>
      <Toaster toasterId={toasterId} position="top-end" offset={{ vertical: 48 }} limit={4} />
      {current && o && (
        <Dialog
          key={current.id}
          open
          modalType="alert"
          onOpenChange={(_, d) => {
            if (!d.open) settle(false);
          }}
        >
          <DialogSurface style={{ maxWidth: 480 }}>
            <DialogBody>
              <DialogTitle>{o.title ?? 'Are you sure?'}</DialogTitle>
              {o.message != null && (
                <DialogContent className={s.message}>{o.message}</DialogContent>
              )}
              <DialogActions>
                <Button appearance="secondary" onClick={() => settle(false)}>
                  {o.cancelLabel ?? 'Cancel'}
                </Button>
                <Button
                  appearance="primary"
                  className={o.destructive ? s.destructive : undefined}
                  onClick={() => settle(true)}
                >
                  {o.confirmLabel ?? (o.destructive ? 'Delete' : 'OK')}
                </Button>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>
      )}
    </>
  );
}
