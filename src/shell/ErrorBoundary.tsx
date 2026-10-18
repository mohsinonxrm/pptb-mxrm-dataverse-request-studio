import { Component, Fragment, useState, type ErrorInfo, type ReactNode } from 'react';
import {
  Button,
  MessageBar,
  MessageBarActions,
  MessageBarBody,
  MessageBarTitle,
  tokens,
} from '@fluentui/react-components';
import { ArrowClockwise20Regular, Checkmark20Regular, Copy20Regular } from '@fluentui/react-icons';
import { copyToClipboardSafe } from '../host/pptbBridge';

// ── Error boundaries (#19) ────────────────────────────────────────────
//
// A render error anywhere below the root used to unmount the whole DRS
// shell (blank tool, no way back short of reopening it). Two layers now
// contain it:
//
//   • scope="mode" — around the routed mode in App's ModeRouter. The header
//     and request-type picker stay alive, so the user can switch modes.
//   • scope="pane" — around the Builder / Code / Results tab body in
//     MainTabs. A crash in the Results grid or the code generator leaves the
//     sidebar, URL bar and the other tabs (and the mode's state) intact.
//
// The fallback never uses native alert/confirm (they freeze PPTB, #111).
// "Reload" remounts the children via a key bump; a change of `resetKey`
// (another mode, another tab) clears the error too.

type Scope = 'mode' | 'pane';

interface Props {
  scope: Scope;
  /** What failed, e.g. "Retrieve Multiple" or "Results". */
  label: string;
  /** Clears a caught error when it changes (mode id, tab id, …). */
  resetKey: unknown;
  children: ReactNode;
}

interface State {
  error: Error | null;
  componentStack: string | null;
  /** Bumped by "Reload" to remount the children from scratch. */
  attempt: number;
}

export class ErrorBoundary extends Component<Props, State> {
  state: State = { error: null, componentStack: null, attempt: 0 };

  static getDerivedStateFromError(error: unknown): Partial<State> {
    return { error: error instanceof Error ? error : new Error(String(error)) };
  }

  componentDidCatch(error: unknown, info: ErrorInfo) {
    this.setState({ componentStack: info.componentStack ?? null });
    console.error(`[DRS] ${this.props.scope} "${this.props.label}" crashed`, error, info);
  }

  componentDidUpdate(prev: Props) {
    if (this.state.error && !Object.is(prev.resetKey, this.props.resetKey)) {
      this.setState({ error: null, componentStack: null });
    }
  }

  private reload = () =>
    this.setState((s) => ({ error: null, componentStack: null, attempt: s.attempt + 1 }));

  render() {
    const { error, componentStack, attempt } = this.state;
    if (!error) return <Fragment key={attempt}>{this.props.children}</Fragment>;
    return (
      <ErrorFallback
        scope={this.props.scope}
        label={this.props.label}
        error={error}
        componentStack={componentStack}
        onReload={this.reload}
      />
    );
  }
}

function errorDetails(scope: Scope, label: string, error: Error, componentStack: string | null) {
  return [
    `Dataverse Request Studio — ${scope === 'mode' ? 'mode' : 'pane'} error`,
    `Where: ${label}`,
    `When: ${new Date().toISOString()}`,
    `User agent: ${typeof navigator !== 'undefined' ? navigator.userAgent : 'n/a'}`,
    '',
    `${error.name}: ${error.message}`,
    '',
    'Stack:',
    error.stack ?? '(none)',
    '',
    'Component stack:',
    componentStack?.trim() || '(none)',
  ].join('\n');
}

function ErrorFallback({
  scope,
  label,
  error,
  componentStack,
  onReload,
}: {
  scope: Scope;
  label: string;
  error: Error;
  componentStack: string | null;
  onReload: () => void;
}) {
  const [copy, setCopy] = useState<'idle' | 'copied' | 'failed'>('idle');
  const onCopy = async () => {
    try {
      await copyToClipboardSafe(errorDetails(scope, label, error, componentStack));
      setCopy('copied');
    } catch {
      setCopy('failed');
    }
  };
  const isMode = scope === 'mode';

  return (
    <div
      role="alert"
      style={{
        // Mode scope fills the frame's body row; pane scope sits inside the
        // tab body, which already has padding.
        padding: isMode ? '24px' : 0,
        minHeight: 0,
        minWidth: 0,
        overflow: 'auto',
        backgroundColor: isMode ? tokens.colorNeutralBackground2 : undefined,
      }}
    >
      <MessageBar layout="multiline" intent="error" style={{ maxWidth: 820 }}>
        <MessageBarBody>
          <MessageBarTitle>
            {isMode ? `The ${label} mode hit an error` : `The ${label} pane hit an error`}
          </MessageBarTitle>
          <div>
            {isMode
              ? 'The rest of DRS still works — reload this mode or pick another request type from the header.'
              : 'The rest of this mode still works — reload this pane or switch to another tab.'}
          </div>
          <code
            style={{
              display: 'block',
              marginTop: 8,
              whiteSpace: 'pre-wrap',
              wordBreak: 'break-word',
              fontFamily: tokens.fontFamilyMonospace,
              fontSize: tokens.fontSizeBase200,
            }}
          >
            {error.name}: {error.message}
          </code>
          {copy === 'failed' && (
            <div style={{ marginTop: 6, color: tokens.colorPaletteRedForeground1 }}>
              Couldn't copy to the clipboard. The details are also in the developer console.
            </div>
          )}
        </MessageBarBody>
        <MessageBarActions>
          <Button appearance="primary" icon={<ArrowClockwise20Regular />} onClick={onReload}>
            {isMode ? 'Reload this mode' : 'Reload this pane'}
          </Button>
          <Button
            icon={copy === 'copied' ? <Checkmark20Regular /> : <Copy20Regular />}
            onClick={onCopy}
          >
            {copy === 'copied' ? 'Copied' : 'Copy error details'}
          </Button>
        </MessageBarActions>
      </MessageBar>
    </div>
  );
}
