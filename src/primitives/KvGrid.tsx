// KvGrid — shared 2-column key/value grid used by every "detail" pane.
//
// Per design-review CS-2: at least six places in the codebase render the
// same `gridTemplateColumns: '160px 1fr'` pattern (Workflow detail card,
// Delete cascade summary, ManageImage column metadata, NavPropertyPicker
// summary, FunctionUrlPreview pairs, ResponsePropertiesCard rows). Each
// reinvented the same grid with slightly different padding, fonts, and
// border colors.
//
// This primitive is the canonical form. Callers pass `Array<{ k, v, mono? }>`
// and get consistent typography + spacing + dark-mode-correct colors.

import { mergeClasses, tokens } from '@fluentui/react-components';
import { useStudioStyles } from './styles';
import type { ReactNode } from 'react';

export interface KvRow {
  /** Label cell — typically a short noun phrase. */
  k: ReactNode;
  /** Value cell — text, code, badges, or any ReactNode. */
  v: ReactNode;
  /** When true, the value cell uses the monospace stack at 10px. */
  mono?: boolean;
}

export interface KvGridProps {
  rows: KvRow[];
  /** Per-grid key column width. Defaults to 160px. */
  keyWidth?: number | string;
  /** Optional className passthrough for wrapper-level styling. */
  className?: string;
}

export function KvGrid({ rows, keyWidth, className }: KvGridProps) {
  const s = useStudioStyles();
  const style =
    keyWidth != null
      ? { gridTemplateColumns: `${typeof keyWidth === 'number' ? `${keyWidth}px` : keyWidth} 1fr` }
      : undefined;
  return (
    <div className={mergeClasses(s.kvGrid, className)} style={style}>
      {rows.map((r, i) => (
        <KvPair key={i} k={r.k} v={r.v} mono={r.mono} />
      ))}
    </div>
  );
}

function KvPair({ k, v, mono }: KvRow) {
  const s = useStudioStyles();
  return (
    <>
      <span className={s.kvKey}>{k}</span>
      <span className={mono ? s.kvValMono : s.kvVal}>
        {/* Empty/falsy values get a faded em-dash so the grid stays aligned. */}
        {v === null || v === undefined || v === '' ? (
          <em style={{ color: tokens.colorNeutralForeground4 }}>—</em>
        ) : (
          v
        )}
      </span>
    </>
  );
}
