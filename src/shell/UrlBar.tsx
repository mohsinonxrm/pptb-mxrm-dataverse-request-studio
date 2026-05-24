import { Button, Tooltip, tokens, Caption1, Spinner } from '@fluentui/react-components';
import {
  Copy20Regular, Play20Filled, Open20Regular,
} from '@fluentui/react-icons';
import type { FC } from 'react';
import { useEffect } from 'react';
import { useStudioStyles } from '../primitives/styles';
import { MethodPill } from '../primitives/MethodPill';
import { AdvisoryDrawer } from '../primitives/AdvisoryDrawer';
import type { Advisory } from '../primitives/advisories';
import { ENV } from '../mock/environment';
import { UrlTokens } from './UrlTokens';
import type { HttpMethod } from '../registry/requestTypes';

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
  onExecute: () => void;
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
}

export function UrlBar({ method, url, disabledReason, executeVerb, executeIcon, onExecute, loading, advisories, onAdvisoryFocus }: UrlBarProps) {
  const ExecIcon = executeIcon ?? Play20Filled;
  const s = useStudioStyles();
  const fullUrl = `https://${ENV.host}${url}`;

  // Ctrl+Enter shortcut
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if ((e.metaKey || e.ctrlKey) && e.key === 'Enter') {
        if (!disabledReason) onExecute();
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [disabledReason, onExecute]);

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
          onClick={() => navigator.clipboard?.writeText(fullUrl)}
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
      <Caption1 style={{
        fontSize: 11, color: tokens.colorNeutralForeground3,
        fontFamily: tokens.fontFamilyMonospace,
      }}>
        {(url.length / 1024).toFixed(1)} KB
      </Caption1>
      {advisories && advisories.length > 0 && (
        <AdvisoryDrawer advisories={advisories} onFocusNode={onAdvisoryFocus} />
      )}
      <Tooltip
        content={disabledReason ?? `${executeVerb} (Ctrl+Enter)`}
        relationship={disabledReason ? 'description' : 'label'}
      >
        <Button
          icon={loading ? <Spinner size="tiny" /> : <ExecIcon />}
          appearance="primary"
          size="medium"
          onClick={onExecute}
          disabled={!!disabledReason || loading}
        >
          {executeVerb}
        </Button>
      </Tooltip>
    </footer>
  );
}
