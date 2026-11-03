import { tokens } from '@fluentui/react-components';
import { useStudioStyles } from './styles';
import { statusTextFor } from '../engine/dataverseExecutor';

export type StatusKind = 'success' | 'danger' | 'warning' | 'info';

export function StatusPill({
  status,
  code,
  ms,
  size,
}: {
  status: StatusKind;
  code: number | string;
  ms?: number;
  size?: number;
}) {
  const s = useStudioStyles();
  // Background1 tints: Foreground1 on them is at least 4.5:1 in light and
  // dark themes (on Background2 it fell to 3.5–4.3:1, #387). The border keeps
  // the paler pill visible on white.
  const colors: Record<StatusKind, { bg: string; fg: string; border: string }> = {
    success: {
      bg: tokens.colorPaletteGreenBackground1,
      fg: tokens.colorPaletteGreenForeground1,
      border: tokens.colorPaletteGreenBorder1,
    },
    danger: {
      bg: tokens.colorPaletteRedBackground1,
      fg: tokens.colorPaletteRedForeground1,
      border: tokens.colorPaletteRedBorder1,
    },
    warning: {
      bg: tokens.colorPaletteYellowBackground1,
      fg: tokens.colorPaletteYellowForeground1,
      border: tokens.colorPaletteYellowBorder1,
    },
    info: {
      bg: tokens.colorNeutralBackground3,
      fg: tokens.colorNeutralForeground2,
      border: tokens.colorNeutralStroke2,
    },
  };
  const c = colors[status];
  // Success badges carry the reason phrase from the shared status map, so a
  // create reads "201 Created" and a 204 "204 No Content" (#123).
  const label =
    status === 'success'
      ? `${code} ${(typeof code === 'number' && statusTextFor(code)) || 'OK'}`
      : `${code}`;
  return (
    <span
      className={s.statusPill}
      style={{
        backgroundColor: c.bg,
        color: c.fg,
        border: `1px solid ${c.border}`,
        fontSize: size ?? 11,
      }}
    >
      {label}
      {/* Not dimmed: at 70% opacity it fell to 2.5:1 (#387) */}
      {ms !== undefined && <span style={{ fontWeight: 400 }}>· {ms} ms</span>}
    </span>
  );
}
