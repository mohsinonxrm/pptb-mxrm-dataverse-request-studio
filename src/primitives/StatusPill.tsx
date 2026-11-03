import {
  Badge,
  makeStyles,
  mergeClasses,
  tokens,
  type BadgeProps,
} from '@fluentui/react-components';
import { Checkmark16Filled } from '@fluentui/react-icons';
import { statusTextFor } from '../engine/dataverseExecutor';

export type StatusKind = 'success' | 'danger' | 'warning' | 'info';

// The response status: a Fluent v9 filled Badge, as the v2.2 design maps it
// (v9-coverage.md: `<Badge appearance="filled" color="success">`), with the
// check icon on success (anatomy.md: status pills, 14–16px Filled) (#387).
const COLOR: Record<StatusKind, BadgeProps['color']> = {
  success: 'success',
  danger: 'danger',
  warning: 'warning',
  info: 'informative',
};

const useStyles = makeStyles({
  badge: {
    fontWeight: tokens.fontWeightSemibold,
    fontVariantNumeric: 'tabular-nums',
    whiteSpace: 'nowrap',
  },
  // Filled success / danger draw on colorNeutralForegroundOnBrand, which the
  // mode cascade (index.html, [data-mode-group]) turns black in dark. Their
  // backgrounds don't change with the theme, so the text stays white.
  onStatic: { color: tokens.colorNeutralForegroundStaticInverted },
  ms: { fontWeight: tokens.fontWeightRegular },
});

export function StatusPill({
  status,
  code,
  ms,
}: {
  status: StatusKind;
  code: number | string;
  ms?: number;
}) {
  const s = useStyles();
  // Success badges carry the reason phrase from the shared status map, so a
  // create reads "201 Created" and a 204 "204 No Content" (#123).
  const label =
    status === 'success'
      ? `${code} ${(typeof code === 'number' && statusTextFor(code)) || 'OK'}`
      : `${code}`;
  return (
    <Badge
      appearance="filled"
      color={COLOR[status]}
      shape="circular"
      size="large"
      icon={status === 'success' ? <Checkmark16Filled /> : undefined}
      className={mergeClasses(s.badge, (status === 'success' || status === 'danger') && s.onStatic)}
    >
      {label}
      {ms !== undefined && <span className={s.ms}>&nbsp;· {ms} ms</span>}
    </Badge>
  );
}
