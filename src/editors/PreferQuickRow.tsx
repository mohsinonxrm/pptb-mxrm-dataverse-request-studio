// The Prefer toggles under the URL bar for write types (#299): the options
// that apply to a create / update / upsert, and one line on what the
// current choice returns. It edits the same PreferSpec as the sidebar's
// Prefer pane, which stays the full editor.
//
// Prefer is host-owned in ToolBox (see PreferEditor): DRS doesn't send it on
// writes, so there's no header-consent prompt and no page-size gate here.
// Inside ToolBox the info icon says what Execute returns regardless; the
// toggles still set the request shown in the Code tab.

import { Caption1, Switch, Tooltip, makeStyles, tokens } from '@fluentui/react-components';
import { Info16Regular } from '@fluentui/react-icons';
import { useHostSession } from '../host/HostContext';
import type { PreferSpec } from './PreferEditor';

export type PreferWriteOperation = 'create' | 'update' | 'upsert';

/** What the current choice returns, in one line. */
export function preferEffectLine(op: PreferWriteOperation, spec: PreferSpec): string {
  if (!spec.returnRepresentation) {
    return op === 'create'
      ? "Returns 204 No Content; the new record's id is in OData-EntityId."
      : 'Returns 204 No Content.';
  }
  const formatted = spec.allAnnotations || spec.formattedValues ? ', with formatted values' : '';
  return op === 'create'
    ? `Returns 201 Created with the new record${formatted}.`
    : `Returns 200 OK with the ${op === 'upsert' ? 'created or updated' : 'updated'} record${formatted}.`;
}

/** What ToolBox's Execute returns, whatever the toggles say (#108). */
const HOST_NOTE: Record<PreferWriteOperation, string> = {
  create:
    'In ToolBox, Execute always returns the new row: ToolBox sets Prefer itself. These toggles set the request in the Code tab.',
  update:
    "In ToolBox, Execute doesn't return the row: ToolBox sets Prefer itself. These toggles set the request in the Code tab.",
  upsert:
    "In ToolBox, Execute doesn't return the row: ToolBox sets Prefer itself. These toggles set the request in the Code tab.",
};

const useStyles = makeStyles({
  row: {
    display: 'flex',
    alignItems: 'center',
    flexWrap: 'wrap',
    columnGap: tokens.spacingHorizontalM,
    minHeight: '32px',
    paddingLeft: '12px',
    paddingRight: '12px',
    backgroundColor: tokens.colorNeutralBackground1,
    borderTop: `1px solid ${tokens.colorNeutralStroke3}`,
    minWidth: 0,
  },
  label: { color: tokens.colorNeutralForeground3, fontWeight: tokens.fontWeightSemibold },
  effect: {
    color: tokens.colorNeutralForeground2,
    minWidth: 0,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  info: { color: tokens.colorNeutralForeground3, display: 'inline-flex', flexShrink: 0 },
});

export function PreferQuickRow({
  operation,
  spec,
  setSpec,
}: {
  operation: PreferWriteOperation;
  spec: PreferSpec;
  setSpec: (s: PreferSpec) => void;
}) {
  const s = useStyles();
  const host = useHostSession();
  const effect = preferEffectLine(operation, spec);
  return (
    <div className={s.row} role="group" aria-label="Prefer">
      <Caption1 className={s.label}>Prefer</Caption1>
      <Switch
        label="return=representation"
        checked={spec.returnRepresentation}
        onChange={(_, d) => setSpec({ ...spec, returnRepresentation: d.checked })}
      />
      <Switch
        label="Formatted values"
        checked={spec.allAnnotations || spec.formattedValues}
        disabled={!spec.returnRepresentation || spec.allAnnotations}
        onChange={(_, d) => setSpec({ ...spec, formattedValues: d.checked })}
      />
      <Caption1 className={s.effect} title={effect}>
        {effect}
      </Caption1>
      {host.embedded && (
        <Tooltip content={HOST_NOTE[operation]} relationship="description">
          <span
            className={s.info}
            role="img"
            tabIndex={0}
            aria-label="What Execute returns in ToolBox"
          >
            <Info16Regular />
          </span>
        </Tooltip>
      )}
    </div>
  );
}
