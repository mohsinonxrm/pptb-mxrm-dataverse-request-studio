// The loaded saved request's name in the header, next to the request-type
// chip, with an "Unsaved" tag while the request has changes (#298). Click
// the name to rename it; Enter (or leaving the field) saves, Esc cancels.
// A request that was never saved shows nothing.

import { useState } from 'react';
import { Badge, Button, Input, Tooltip, makeStyles, tokens } from '@fluentui/react-components';
import { useSavedRequests } from '../state/useSavedRequests';

const useStyles = makeStyles({
  root: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    minWidth: 0,
    flexShrink: 1,
  },
  name: {
    minWidth: 0,
    maxWidth: '240px',
    fontWeight: tokens.fontWeightRegular,
    color: tokens.colorNeutralForeground2,
  },
  // The label truncates in a narrow pane.
  label: { overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' },
  input: { width: '220px' },
  tag: { flexShrink: 0 },
});

export interface SavedNameTagProps {
  /** The saved entry the current request was loaded from or last saved as. */
  lastSavedId?: string;
  /** The request has changes since that save. */
  dirty: boolean;
}

export function SavedNameTag({ lastSavedId, dirty }: SavedNameTagProps) {
  const s = useStyles();
  const { findById, rename } = useSavedRequests();
  const [draft, setDraft] = useState<string | null>(null);
  const entry = lastSavedId ? findById(lastSavedId) : undefined;
  if (!entry) return null;

  const commit = () => {
    const next = draft?.trim();
    if (next && next !== entry.name) rename(entry.id, next);
    setDraft(null);
  };

  return (
    <span className={s.root} data-testid="saved-name">
      {draft === null ? (
        <Tooltip content={`${entry.name} — click to rename`} relationship="description">
          <Button
            appearance="transparent"
            size="small"
            className={s.name}
            aria-label={`Saved request ${entry.name}. Rename`}
            onClick={() => setDraft(entry.name)}
          >
            <span className={s.label}>{entry.name}</span>
          </Button>
        </Tooltip>
      ) : (
        <Input
          appearance="underline"
          size="small"
          className={s.input}
          aria-label="Saved request name"
          autoFocus
          value={draft}
          onChange={(_, d) => setDraft(d.value)}
          onBlur={commit}
          onKeyDown={(e) => {
            if (e.key === 'Enter') commit();
            if (e.key === 'Escape') {
              e.stopPropagation();
              setDraft(null);
            }
          }}
        />
      )}
      {dirty && (
        <Badge appearance="tint" color="informative" size="small" className={s.tag}>
          Unsaved
        </Badge>
      )}
    </span>
  );
}
