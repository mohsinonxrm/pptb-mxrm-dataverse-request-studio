import { Button, Tooltip, tokens, Caption1, Spinner } from '@fluentui/react-components';
import { Copy20Regular, Play20Filled, Open20Regular } from '@fluentui/react-icons';
import type { FC } from 'react';
import { useCallback, useEffect, useRef } from 'react';
import { useStudioStyles } from '../primitives/styles';
import { MethodPill } from '../primitives/MethodPill';
import { AdvisoryDrawer } from '../primitives/AdvisoryDrawer';
import { announce } from '../primitives/announce';
import type { Advisory } from '../primitives/advisories';
import { ENV } from '../mock/environment';
import { UrlTokens } from './UrlTokens';
import type { HttpMethod } from '../registry/requestTypes';
import { getSession } from '../host/pptbBridge';
import { usePersistedSettings } from '../host/usePersistedSettings';
import { needsWriteConfirmation, type WriteConfirmSpec } from '../engine/writeConfirm';
import { confirmWrite } from './confirmWrite';
import { notify } from './dialogs';
import { copyText } from './copy';
import { useOpenAdvisoryNode } from './builderNav';

export interface UrlBarProps {
  method: HttpMethod;
  /** Path relative to host, e.g. /api/data/v9.2/accounts?$select=… */
  url: string;
  /** Reason the Execute button is disabled, if any */
  disabledReason?: string | null;
  /** Verb shown on the Execute button (e.g. "Execute", "Send", "Run"). */
  executeVerb: string;
  /**
   * Icon component for the Execute button. Defaults to Play20Filled.
   * Per-mode overrides:
   *   • Delete → Delete20Filled (trash) — safety affordance
   *   • Action/Function/Custom API/Workflow → Flash20Filled (bolt) — "Run"
   *   • Manage File upload → ArrowUpload20Filled · download → ArrowDownload20Filled
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  executeIcon?: FC<any>;
  onExecute: () => void | Promise<unknown>;
  loading: boolean;

  /**
   * Optional aggregated advisories — wildcards stripped, antipatterns,
   * bypass-header warnings, etc. Renders an AdvisoryDrawer chip between
   * the byte counter and Execute. When empty/undefined, no chip.
   *
   * The mode is responsible for *also* setting disabledReason for any
   * error-severity advisory (use disabledReasonFromAdvisories()).
   */
  advisories?: Advisory[];
  /** Called when the user clicks "Open" on an advisory — focuses that node. */
  onAdvisoryFocus?: (nodeId: string) => void;
  /**
   * Request size shown in the byte counter, when it isn't the URL — e.g. the
   * multipart body of a $batch (#198). Defaults to the URL length.
   */
  bytes?: number;
  /**
   * What the request changes, for the write confirmation (#241). Execute
   * (and Ctrl+Enter) asks first when engine/writeConfirm says so. Omit, or
   * pass null, for reads and for modes that confirm on their own (Bulk
   * Delete Job).
   */
  writeConfirm?: WriteConfirmSpec | null;
}

export function UrlBar({
  method,
  url,
  disabledReason,
  executeVerb,
  executeIcon,
  onExecute,
  loading,
  advisories,
  onAdvisoryFocus,
  bytes,
  writeConfirm,
}: UrlBarProps) {
  const ExecIcon = executeIcon ?? Play20Filled;
  const s = useStudioStyles();
  // Open on an advisory: show the Builder, select the node, focus its field (#294).
  const openAdvisoryNode = useOpenAdvisoryNode(onAdvisoryFocus);
  const fullUrl = `https://${ENV.host}${url}`;
  const [{ productionSafeMode }] = usePersistedSettings();

  // Re-entry guard. `loading` only flips after the mode re-renders, so a
  // held / repeated Ctrl+Enter (or a click in the same frame) could start a
  // second execute — a duplicate write for Create and friends (#47). The ref
  // is set synchronously before `onExecute` runs and released when its
  // promise settles (every mode's onExecute is async), so it never depends
  // on render timing and can't stay stuck. `runtime` never rejects (#102),
  // so a throw here is a bug in the mode's own code; log it rather than
  // leave an unhandled rejection.
  const busyRef = useRef(false);
  // The latest props, read when Execute runs and again after the write
  // confirmation closes: the request can change, or become invalid, while the
  // dialog is open. Reading them through a ref also keeps `execute` (and the
  // Ctrl+Enter listener) from being rebuilt on every render, since modes pass
  // a new `writeConfirm` object each time.
  const latest = useRef({
    loading,
    disabledReason,
    writeConfirm,
    onExecute,
    productionSafeMode,
    executeVerb,
  });
  latest.current = {
    loading,
    disabledReason,
    writeConfirm,
    onExecute,
    productionSafeMode,
    executeVerb,
  };
  const mounted = useRef(true);
  useEffect(() => {
    mounted.current = true;
    return () => {
      mounted.current = false;
    };
  }, []);
  const execute = useCallback(() => {
    const now = latest.current;
    if (busyRef.current || now.loading || now.disabledReason) return;
    busyRef.current = true;
    // Writes ask first (#241). The session is read now, not at render, so the
    // confirmation names the connection the request will actually use. The
    // re-entry guard stays held while the dialog is open.
    const session = getSession();
    const spec = now.writeConfirm;
    const ask = needsWriteConfirmation(spec, session.environment, now.productionSafeMode);
    Promise.resolve(
      ask && spec
        ? confirmWrite(spec, session, {
            productionSafeMode: now.productionSafeMode,
            verb: now.executeVerb,
          })
        : true,
    )
      .then((ok) => {
        if (!ok) return undefined;
        if (!ask) return now.onExecute();
        // Confirmed. Send only if that still holds: the mode is still open,
        // the request is still valid, and the connection is the one the
        // dialog named. Switching connections while it's open would otherwise
        // send the write to an environment the user never confirmed.
        const after = latest.current;
        if (!mounted.current || after.loading || after.disabledReason) return undefined;
        if (getSession().environmentUrl !== session.environmentUrl) {
          notify(
            `The connection changed while the confirmation was open. ${after.executeVerb} again to confirm the new target.`,
            { title: 'Not sent' },
          );
          return undefined;
        }
        return after.onExecute();
      })
      .catch((e: unknown) => {
        console.error('Execute failed', e);
      })
      .finally(() => {
        busyRef.current = false;
      });
  }, []);

  // Ctrl+Enter shortcut
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if ((e.metaKey || e.ctrlKey) && e.key === 'Enter') {
        if (e.repeat) return;
        execute();
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [execute]);

  // Screen-reader feedback that Execute started; ResultsView announces the
  // outcome (#20).
  useEffect(() => {
    if (loading) announce('Running request…');
  }, [loading]);

  return (
    <footer className={s.footer}>
      <MethodPill method={method} />
      <Tooltip content={fullUrl} relationship="description" positioning="above">
        <span className={s.urlText} title={fullUrl}>
          <UrlTokens url={url} host={ENV.host} />
        </span>
      </Tooltip>
      <Tooltip content="Copy URL" relationship="label">
        <Button
          icon={<Copy20Regular />}
          appearance="subtle"
          size="small"
          onClick={() => void copyText(fullUrl, 'the URL')}
        />
      </Tooltip>
      <Tooltip content="Open in new tab" relationship="label">
        <Button
          icon={<Open20Regular />}
          appearance="subtle"
          size="small"
          onClick={() => window.open(fullUrl, '_blank', 'noopener')}
        />
      </Tooltip>
      {/* Byte counter is functional telemetry (not decoration) — 11px@100%
          reads as AA-large without competing with the URL itself; 10px@70%
          opacity is below the legibility floor on standard DPI. */}
      <Caption1
        style={{
          fontSize: 11,
          color: tokens.colorNeutralForeground3,
          fontFamily: tokens.fontFamilyMonospace,
        }}
      >
        {((bytes ?? url.length) / 1024).toFixed(1)} KB
      </Caption1>
      {advisories && advisories.length > 0 && (
        <AdvisoryDrawer advisories={advisories} onFocusNode={openAdvisoryNode} />
      )}
      <Tooltip
        content={disabledReason ?? `${executeVerb} (Ctrl+Enter)`}
        relationship={disabledReason ? 'description' : 'label'}
      >
        <Button
          icon={loading ? <Spinner size="tiny" /> : <ExecIcon />}
          appearance="primary"
          size="medium"
          onClick={execute}
          disabled={!!disabledReason || loading}
        >
          {executeVerb}
        </Button>
      </Tooltip>
    </footer>
  );
}
