import { tokens } from '@fluentui/react-components';
import { useStudioStyles } from './styles';

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
  const colors: Record<StatusKind, { bg: string; fg: string }> = {
    success: { bg: tokens.colorPaletteGreenBackground2, fg: tokens.colorPaletteGreenForeground1 },
    danger: { bg: tokens.colorPaletteRedBackground2, fg: tokens.colorPaletteRedForeground1 },
    warning: { bg: tokens.colorPaletteYellowBackground2, fg: tokens.colorPaletteYellowForeground1 },
    info: { bg: tokens.colorNeutralBackground3, fg: tokens.colorNeutralForeground2 },
  };
  const c = colors[status];
  const label =
    status === 'success'
      ? `${code} OK`
      : status === 'danger'
        ? `${code}`
        : status === 'warning'
          ? `${code}`
          : `${code}`;
  return (
    <span
      className={s.statusPill}
      style={{ backgroundColor: c.bg, color: c.fg, fontSize: size ?? 11 }}
    >
      {label}
      {ms !== undefined && <span style={{ opacity: 0.7 }}>· {ms} ms</span>}
    </span>
  );
}
