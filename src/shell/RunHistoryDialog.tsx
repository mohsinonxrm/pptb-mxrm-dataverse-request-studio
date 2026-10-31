// The run history of one mode (#226): every run kept for this environment,
// searchable, each restorable like a Recent runs row.

import { useMemo, useState } from 'react';
import {
  Button,
  Caption1,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Input,
  tokens,
} from '@fluentui/react-components';
import { Search20Regular } from '@fluentui/react-icons';
import { MethodPill } from '../primitives/MethodPill';
import { useStudioStyles } from '../primitives/styles';
import { usePersistedSettings } from '../host/usePersistedSettings';
import type { RecentRun } from '../state/readState';
import { runDetail, timeAgo } from './runLabels';

/** Runs whose method, URL, status or summary contain every word of `query`. */
export function filterRuns(runs: readonly RecentRun[], query: string): RecentRun[] {
  const words = query.toLowerCase().split(/\s+/).filter(Boolean);
  if (words.length === 0) return [...runs];
  return runs.filter((r) => {
    const text = `${r.method} ${r.url} ${r.status || 'error'} ${r.summary ?? ''}`.toLowerCase();
    return words.every((w) => text.includes(w));
  });
}

export function RunHistoryDialog({
  open,
  onOpenChange,
  modeName,
  runs,
  onRestore,
}: {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  modeName: string;
  runs: readonly RecentRun[];
  /** Restore a run; absent when the mode can't restore runs. */
  onRestore?: (r: RecentRun) => void;
}) {
  const s = useStudioStyles();
  const [{ keepRunHistory }] = usePersistedSettings();
  const [query, setQuery] = useState('');
  const shown = useMemo(() => filterRuns(runs, query), [runs, query]);
  const restorable = (r: RecentRun) => !!onRestore && r.snapshot !== undefined;
  const restore = (r: RecentRun) => {
    onRestore?.(r);
    onOpenChange(false);
  };

  return (
    <Dialog open={open} onOpenChange={(_, d) => onOpenChange(d.open)}>
      <DialogSurface style={{ maxWidth: 720 }}>
        <DialogBody>
          <DialogTitle>{modeName} history</DialogTitle>
          <DialogContent style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
            <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
              {keepRunHistory
                ? 'Runs against this environment, kept across restarts: the request, status, row count and time. Response bodies, credentials and impersonation headers are never kept. Turn it off or clear it in Settings.'
                : 'Request history is off, so these are this session’s runs only. Turn it on in Settings.'}
            </Caption1>
            <Input
              aria-label="Search runs"
              placeholder="Search by URL, status or method…"
              contentBefore={<Search20Regular />}
              value={query}
              onChange={(_, d) => setQuery(d.value)}
            />
            <div
              role="list"
              aria-label={`${modeName} runs`}
              style={{ maxHeight: 420, overflow: 'auto', display: 'flex', flexDirection: 'column' }}
            >
              {shown.length === 0 && (
                <Caption1 style={{ padding: 12, color: tokens.colorNeutralForeground3 }}>
                  {runs.length === 0 ? 'No runs yet.' : 'No runs match the search.'}
                </Caption1>
              )}
              {shown.map((r) => (
                <div
                  key={r.id}
                  role="listitem"
                  className={restorable(r) ? s.recentRow : s.recentRowStatic}
                  tabIndex={restorable(r) ? 0 : undefined}
                  aria-label={
                    restorable(r)
                      ? `Restore run: ${r.method} ${r.url}, status ${r.status || 'error'}`
                      : undefined
                  }
                  title={
                    `${new Date(r.ts).toLocaleString()} · ${r.url}` +
                    (restorable(r) ? '' : ' · This run can’t be restored.')
                  }
                  onClick={() => restorable(r) && restore(r)}
                  onKeyDown={(e) => {
                    if (restorable(r) && (e.key === 'Enter' || e.key === ' ')) {
                      e.preventDefault();
                      restore(r);
                    }
                  }}
                >
                  <MethodPill method={r.method as never} size="sm" />
                  <span style={{ flexShrink: 0, minWidth: 150 }}>
                    {timeAgo(r.ts)} · {r.status || 'error'}
                    {runDetail(r)}
                  </span>
                  <span
                    style={{
                      flexGrow: 1,
                      overflow: 'hidden',
                      textOverflow: 'ellipsis',
                      whiteSpace: 'nowrap',
                      fontFamily: tokens.fontFamilyMonospace,
                      fontSize: 11,
                      color: tokens.colorNeutralForeground3,
                    }}
                  >
                    {r.url}
                  </span>
                </div>
              ))}
            </div>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={() => onOpenChange(false)}>
              Close
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
