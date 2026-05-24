import type { ReactNode } from 'react';
import { useStudioStyles } from './styles';

export function SectionHeader({ children, meta }: { children: ReactNode; meta?: ReactNode }) {
  const s = useStudioStyles();
  return (
    <div className={s.sectionH}>
      <span>{children}</span>
      {/* Meta counter ("3 active", "2 of 5") is 11px — sits comfortably
          below the 12px label without sacrificing visual hierarchy. 9px
          would be below the legibility floor on standard DPI. */}
      {meta && (
        <span style={{ fontWeight: 500, textTransform: 'none', letterSpacing: 0, fontSize: 11 }}>
          · {meta}
        </span>
      )}
    </div>
  );
}
