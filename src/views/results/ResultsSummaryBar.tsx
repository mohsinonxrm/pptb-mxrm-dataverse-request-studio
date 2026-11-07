// One summary line at the top of every results view (#290): status, time,
// size, then for collections the rows loaded so far, the server total
// (`@odata.count`) and whether more pages follow (`@odata.nextLink`).
// The grid footer keeps only what's about the grid itself (filters,
// selection, loading more pages).

import { makeStyles, tokens } from '@fluentui/react-components';
import { StatusPill, type StatusKind } from '../../primitives/StatusPill';

const useStyles = makeStyles({
  bar: {
    display: 'flex',
    alignItems: 'center',
    columnGap: tokens.spacingHorizontalS,
    rowGap: tokens.spacingVerticalXS,
    flexWrap: 'wrap',
    minWidth: 0,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    fontVariantNumeric: 'tabular-nums',
  },
  sep: { color: tokens.colorNeutralForeground4 },
  strong: { color: tokens.colorNeutralForeground1, fontWeight: tokens.fontWeightSemibold },
});

export interface ResultsSummaryBarProps {
  /** HTTP status; 0 when the host didn't report one. */
  code: number;
  ms?: number;
  bytes?: number;
  /** Overrides the status colour (an error result reports 'danger'). */
  kind?: StatusKind;
  /** What the pill shows when `code` is 0. */
  codeLabel?: string;
  /** Rows loaded so far, for a collection (pages append as they load). */
  rows?: number;
  /** `@odata.count`, when the request asked for it. */
  total?: number;
  /** `@odata.nextLink` is present. */
  more?: boolean;
}

export const statusKindOf = (code: number): StatusKind =>
  code >= 400 || code === 0 ? 'danger' : code >= 300 ? 'warning' : 'success';

/** "0.9 KB", "1.2 MB". */
export function formatBytes(bytes: number): string {
  return bytes >= 1024 * 1024
    ? `${(bytes / (1024 * 1024)).toFixed(1)} MB`
    : `${(bytes / 1024).toFixed(1)} KB`;
}

export function ResultsSummaryBar({
  code,
  ms,
  bytes,
  kind,
  codeLabel,
  rows,
  total,
  more,
}: ResultsSummaryBarProps) {
  const s = useStyles();
  const parts: React.ReactNode[] = [];
  if (bytes !== undefined) parts.push(<span key="size">{formatBytes(bytes)}</span>);
  if (rows !== undefined) {
    parts.push(
      <span key="rows">
        <span className={s.strong}>{rows.toLocaleString()}</span> row{rows === 1 ? '' : 's'}
        {more ? ' loaded' : ''}
      </span>,
    );
  }
  if (total !== undefined) {
    parts.push(
      <span key="total">
        total <span className={s.strong}>{total.toLocaleString()}</span>
      </span>,
    );
  }
  if (more) parts.push(<span key="more">more pages</span>);
  return (
    <div className={s.bar} data-testid="results-summary">
      <StatusPill status={kind ?? statusKindOf(code)} code={code || codeLabel || 'Error'} ms={ms} />
      {parts.map((p, i) => (
        <span key={i} style={{ display: 'contents' }}>
          {i > 0 && (
            <span className={s.sep} aria-hidden>
              ·
            </span>
          )}
          {p}
        </span>
      ))}
    </div>
  );
}
