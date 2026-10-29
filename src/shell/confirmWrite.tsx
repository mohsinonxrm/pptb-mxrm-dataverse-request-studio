// The confirmation shown before a write is sent (#241). It names the target:
// the environment (connection name, URL and type, with Production
// highlighted), the operation, the table and the scope. The rules for when it
// shows, and when a word must be typed, are in engine/writeConfirm.ts.

import { Badge, Caption1, tokens } from '@fluentui/react-components';
import { confirmDialog } from './dialogs';
import {
  confirmWordOf,
  describeWriteScope,
  isProductionEnvironment,
  requiresTypedConfirmation,
  type WriteConfirmSpec,
} from '../engine/writeConfirm';

export interface WriteTarget {
  /** Connection name in ToolBox, e.g. "Contoso Prod". */
  connectionName: string;
  /** Environment URL, e.g. "https://contoso.crm.dynamics.com". */
  environmentUrl: string;
  /** "Dev" / "Test" / "UAT" / "Production", or empty. */
  environment: string;
}

const row = { display: 'contents' } as const;
const label = { color: tokens.colorNeutralForeground3 } as const;

/** The environment as a badge: Production in red, the others neutral. */
export function EnvironmentBadge({ environment }: { environment: string }) {
  if (!environment) return null;
  const prod = isProductionEnvironment(environment);
  return (
    <Badge
      appearance={prod ? 'filled' : 'tint'}
      color={prod ? 'danger' : 'informative'}
      size="small"
    >
      {environment}
    </Badge>
  );
}

function TargetSummary({ spec, target }: { spec: WriteConfirmSpec; target: WriteTarget }) {
  const host = target.environmentUrl.replace(/^https?:\/\//, '').replace(/\/$/, '');
  return (
    <div
      style={{
        display: 'grid',
        gridTemplateColumns: 'auto 1fr',
        columnGap: tokens.spacingHorizontalM,
        rowGap: tokens.spacingVerticalXS,
        padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
        background: tokens.colorNeutralBackground3,
        borderRadius: tokens.borderRadiusMedium,
      }}
    >
      <div style={row}>
        <span style={label}>Environment</span>
        <span style={{ display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 8 }}>
          <strong>{target.connectionName || host || '(no connection)'}</strong>
          <EnvironmentBadge environment={target.environment} />
          {host && target.connectionName && (
            <Caption1 style={{ color: tokens.colorNeutralForeground3, wordBreak: 'break-all' }}>
              {host}
            </Caption1>
          )}
        </span>
      </div>
      <div style={row}>
        <span style={label}>Operation</span>
        <span>{spec.operation}</span>
      </div>
      {spec.table && (
        <div style={row}>
          <span style={label}>Table</span>
          <code style={{ fontFamily: tokens.fontFamilyMonospace }}>{spec.table}</code>
        </div>
      )}
      <div style={row}>
        <span style={label}>Scope</span>
        <span>{describeWriteScope(spec.scope)}</span>
      </div>
    </div>
  );
}

/**
 * Ask before sending a write. Resolves true to go ahead. The caller decides
 * whether to ask at all (needsWriteConfirmation).
 */
export function confirmWrite(
  spec: WriteConfirmSpec,
  target: WriteTarget,
  options: { productionSafeMode: boolean; verb: string },
): Promise<boolean> {
  const prod = isProductionEnvironment(target.environment);
  const typed = requiresTypedConfirmation(spec, target.environment, options.productionSafeMode);
  return confirmDialog({
    title: `Confirm: ${spec.operation}`,
    wide: true,
    destructive: spec.highImpact || prod,
    confirmLabel: options.verb,
    typeToConfirm: typed ? confirmWordOf(spec) : undefined,
    message: (
      <>
        <TargetSummary spec={spec} target={target} />
        <p style={{ margin: `${tokens.spacingVerticalM} 0 0` }}>
          {prod
            ? 'This is a Production environment. The request runs as soon as you confirm, and it changes data there.'
            : spec.highImpact
              ? "This request changes data and can't be undone from here."
              : 'This request changes data.'}
        </p>
      </>
    ),
  });
}
