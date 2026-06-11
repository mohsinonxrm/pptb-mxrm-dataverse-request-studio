import type { FC, ReactNode } from 'react';
import { useStudioStyles } from '../primitives/styles';
import { groupColorVar } from '../theme/theme';
import type { RequestGroup } from '../registry/requestTypes';

export function PaneHead({
  icon: Icon,
  title,
  sub,
  group = 'read',
  children,
}: {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  icon?: FC<any>;
  title: string;
  sub?: ReactNode;
  group?: RequestGroup;
  children?: ReactNode;
}) {
  const s = useStudioStyles();
  return (
    <div className={s.paneHead}>
      {Icon && <Icon style={{ color: groupColorVar(group), width: 22, height: 22 }} />}
      <span className={s.paneTitle}>{title}</span>
      {sub && <span className={s.paneSub}>{sub}</span>}
      <span style={{ flexGrow: 1 }} />
      {children}
    </div>
  );
}
