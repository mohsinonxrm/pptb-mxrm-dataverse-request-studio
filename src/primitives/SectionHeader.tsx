import type { ReactNode } from 'react';
import { useStudioStyles } from './styles';

export function SectionHeader({
  children,
  meta,
  id,
}: {
  children: ReactNode;
  meta?: ReactNode;
  /** Id on the label, for a list it names (#243). */
  id?: string;
}) {
  const s = useStudioStyles();
  return (
    // A heading (#243): screen-reader users can jump between sidebar sections.
    <div className={s.sectionH} role="heading" aria-level={2}>
      <span id={id}>{children}</span>
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
