// Saved-requests UX: two surfaces in one file because they share state.
//
//   <SaveButton>         — dirty-aware button in the URL bar. Opens the
//                           save dialog with an auto-suggested name.
//   <SaveDialog>         — name input + Save / Cancel. If the name
//                           collides with an existing entry, shows an
//                           overwrite confirm.
//   <SavedLibraryButton> — opens a popover listing all saved entries
//                           with click-to-load + rename + delete.
//
// Library lives in localStorage (see savedRequests.ts). Per-mode load
// routing is the caller's responsibility — the dialog just hands back
// the SavedRequest; the mode decides what to do with `state`.

import { useEffect, useMemo, useState } from 'react';
import {
  Dialog, DialogTrigger, DialogSurface, DialogTitle, DialogBody, DialogActions, DialogContent,
  Popover, PopoverTrigger, PopoverSurface,
  Button, Input, Caption1, Badge, tokens, Tooltip,
  MessageBar, MessageBarBody, MessageBarTitle,
  mergeClasses, makeStyles,
} from '@fluentui/react-components';
import {
  Save20Regular, Save20Filled, BookmarkMultiple20Regular,
  Delete20Regular, Edit20Regular, Dismiss20Regular, Open16Regular,
  Checkmark16Filled,
} from '@fluentui/react-icons';
import {
  autoSuggestName, newSavedId, getOrgScope, tableNameFromState,
  type SavedRequest, type SavedModeId,
} from '../state/savedRequests';
import { useSavedRequests } from '../state/useSavedRequests';

const useStyles = makeStyles({
  libraryList: {
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
    maxHeight: '420px',
    overflowY: 'auto',
    minWidth: '380px',
  },
  libraryRow: {
    display: 'grid',
    gridTemplateColumns: 'auto 1fr auto',
    gap: '8px',
    alignItems: 'center',
    padding: '8px 10px',
    borderRadius: tokens.borderRadiusMedium,
    cursor: 'pointer',
    borderTopWidth: '1px',
    borderRightWidth: '1px',
    borderBottomWidth: '1px',
    borderLeftWidth: '1px',
    borderTopStyle: 'solid',
    borderRightStyle: 'solid',
    borderBottomStyle: 'solid',
    borderLeftStyle: 'solid',
    borderTopColor: 'transparent',
    borderRightColor: 'transparent',
    borderBottomColor: 'transparent',
    borderLeftColor: 'transparent',
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground1Hover,
      borderTopColor: tokens.colorNeutralStroke2,
      borderRightColor: tokens.colorNeutralStroke2,
      borderBottomColor: tokens.colorNeutralStroke2,
      borderLeftColor: tokens.colorNeutralStroke2,
    },
  },
  libraryRowName: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  libraryRowMeta: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    fontFamily: tokens.fontFamilyMonospace,
  },
  libraryEmpty: {
    padding: '32px 16px',
    textAlign: 'center',
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
});

// ═════════════════════════════════════════════════════════════════════
//  Save button — used in the UrlBar
// ═════════════════════════════════════════════════════════════════════

export interface SaveButtonProps {
  /** Current snapshot to save. Type-erased — the SaveButton just hands
   *  it to localStorage. The mode owns the actual shape. */
  state: unknown;
  /** Mode id — written into the saved entry for library filtering / load routing. */
  modeId: SavedModeId;
  /** True when the snapshot has diverged from the last-saved state OR no save exists yet.
   *  Drives the button's enabled state (Fluent 2 standard pattern: Save is enabled
   *  exactly when there's something new to persist). */
  dirty: boolean;
  /** If the user previously saved this request, the existing entry's id. Lets
   *  Save (no "As") overwrite that entry without prompting. */
  lastSavedId?: string;
  /** Called after a successful save with the saved entry. The mode uses it to
   *  remember `lastSavedId` for subsequent overwrites + to refresh dirty state. */
  onSaved: (saved: SavedRequest) => void;
}

export function SaveButton({ state, modeId, dirty, lastSavedId, onSaved }: SaveButtonProps) {
  const [open, setOpen] = useState(false);
  const [name, setName] = useState('');
  const [confirmOverwrite, setConfirmOverwrite] = useState(false);
  const { save, findById, findByName, error } = useSavedRequests();

  // Auto-suggest name whenever the dialog opens. If the user previously
  // saved this request, default to the existing name so a quick re-save
  // is one click + Enter.
  useEffect(() => {
    if (!open) return;
    if (lastSavedId) {
      const existing = findById(lastSavedId);
      setName(existing?.name ?? autoSuggestName(modeId, state));
    } else {
      setName(autoSuggestName(modeId, state));
    }
    setConfirmOverwrite(false);
  }, [open, state, modeId, lastSavedId, findById]);

  const doSave = (overwriteId?: string) => {
    const trimmed = name.trim();
    if (!trimmed) return;
    const id = overwriteId ?? newSavedId();
    const entry: SavedRequest = {
      id,
      name: trimmed,
      modeId,
      orgScope: getOrgScope(),
      state,
      savedAt: Date.now(),
    };
    const r = save(entry);
    if (!r.ok) return; // error surfaced via useSavedRequests().error
    onSaved(entry);
    setOpen(false);
  };

  const onClickSave = () => {
    const trimmed = name.trim();
    if (!trimmed) return;
    // Overwrite paths:
    //  • Same id (the entry we previously saved as): silent overwrite.
    //  • Same name, different id (collision): prompt confirm.
    if (lastSavedId && findById(lastSavedId)?.name === trimmed) {
      doSave(lastSavedId);
      return;
    }
    const byName = findByName(trimmed);
    if (byName) {
      // Collision with a DIFFERENT entry → confirm overwrite or rename.
      setConfirmOverwrite(true);
      return;
    }
    doSave();
  };

  // Render: the trigger button itself (in the URL bar) + the dialog content.
  return (
    <Dialog open={open} onOpenChange={(_, d) => setOpen(d.open)}>
      <DialogTrigger disableButtonEnhancement>
        <Tooltip
          content={dirty
            ? 'Save current request (unsaved changes)'
            : 'No changes since last save'}
          relationship="label"
        >
          <Button
            appearance={dirty ? 'primary' : 'subtle'}
            icon={dirty ? <Save20Filled /> : <Save20Regular />}
            disabled={!dirty}
            aria-label="Save request"
          >
            Save
          </Button>
        </Tooltip>
      </DialogTrigger>
      <DialogSurface style={{ maxWidth: 520 }}>
        <DialogBody>
          <DialogTitle
            action={
              <DialogTrigger action="close">
                <Button appearance="subtle" icon={<Dismiss20Regular />} aria-label="close" />
              </DialogTrigger>
            }
          >
            Save request
          </DialogTitle>
          <DialogContent>
            <Caption1 style={{ display: 'block', marginBottom: 6, color: tokens.colorNeutralForeground2 }}>
              Saved requests live in this browser&apos;s storage, scoped to{' '}
              <code style={{ fontFamily: tokens.fontFamilyMonospace }}>{getOrgScope()}</code>.
              Other Dataverse connections have their own saved-request lists. Name auto-suggested
              from the current entity, select, and filter counts.
            </Caption1>
            <Input
              value={name}
              onChange={(_, d) => { setName(d.value); setConfirmOverwrite(false); }}
              placeholder="Untitled request"
              style={{ width: '100%' }}
              autoFocus
              onKeyDown={(e) => {
                if (e.key === 'Enter' && !confirmOverwrite) {
                  e.preventDefault();
                  onClickSave();
                }
              }}
            />
            {confirmOverwrite && (
              <MessageBar layout="multiline" intent="warning" style={{ marginTop: 12 }}>
                <MessageBarBody>
                  <MessageBarTitle>Name in use</MessageBarTitle>
                  A saved request already uses <code>{name.trim()}</code>. Save again to overwrite
                  it, or change the name above and Save.
                </MessageBarBody>
              </MessageBar>
            )}
            {error && (
              <MessageBar layout="multiline" intent="error" style={{ marginTop: 12 }}>
                <MessageBarBody>
                  <MessageBarTitle>Couldn&apos;t save</MessageBarTitle>
                  <code>{error}</code>
                </MessageBarBody>
              </MessageBar>
            )}
          </DialogContent>
          <DialogActions>
            <DialogTrigger action="close" disableButtonEnhancement>
              <Button appearance="subtle">Cancel</Button>
            </DialogTrigger>
            <Button
              appearance="primary"
              icon={<Save20Filled />}
              onClick={() => {
                if (confirmOverwrite) {
                  const existing = findByName(name.trim());
                  if (existing) doSave(existing.id);
                } else {
                  onClickSave();
                }
              }}
              disabled={!name.trim()}
            >
              {confirmOverwrite ? 'Overwrite' : 'Save'}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

// ═════════════════════════════════════════════════════════════════════
//  Saved library — popover triggered from the UrlBar
// ═════════════════════════════════════════════════════════════════════

export interface SavedLibraryButtonProps {
  modeId: SavedModeId;
  /** Called when the user clicks a saved entry. The mode decides how
   *  to apply it (validates against current metadata, then hydrates state). */
  onLoad: (entry: SavedRequest) => void;
  /** Optional — the currently-loaded entry's id, shown with a checkmark badge. */
  currentId?: string;
}

export function SavedLibraryButton({ modeId, onLoad, currentId }: SavedLibraryButtonProps) {
  const s = useStyles();
  const [open, setOpen] = useState(false);
  const [renameId, setRenameId] = useState<string | null>(null);
  const [renameDraft, setRenameDraft] = useState('');
  const { saved, rename, remove } = useSavedRequests();

  // Filter to this mode's entries AND the current org scope. The bucket
  // is already org-scoped via getSavedRequestsKey(), so this filter is a
  // belt-and-suspenders guard — if a corrupt entry slipped in via manual
  // migration or a future bulk-import feature, it stays hidden here.
  // Future modes get their own routing automatically: entries are tagged
  // with modeId at save time.
  const scope = getOrgScope();
  const list = useMemo(
    () => saved
      .filter(e => e.modeId === modeId && (!e.orgScope || e.orgScope === scope))
      .sort((a, b) => b.savedAt - a.savedAt),
    [saved, modeId, scope],
  );

  const onCommitRename = () => {
    if (!renameId) return;
    const trimmed = renameDraft.trim();
    if (trimmed) rename(renameId, trimmed);
    setRenameId(null);
  };

  return (
    <Popover open={open} onOpenChange={(_, d) => { setOpen(d.open); setRenameId(null); }}>
      <PopoverTrigger disableButtonEnhancement>
        <Tooltip content="Open saved requests" relationship="label">
          <Button
            appearance="subtle"
            icon={<BookmarkMultiple20Regular />}
            aria-label="Open saved requests"
          >
            Saved
            {list.length > 0 && (
              <Badge appearance="tint" size="small" style={{ marginLeft: 6 }}>
                {list.length}
              </Badge>
            )}
          </Button>
        </Tooltip>
      </PopoverTrigger>
      <PopoverSurface style={{ padding: 8 }}>
        <Caption1 style={{ display: 'block', padding: '6px 10px', color: tokens.colorNeutralForeground3 }}>
          {list.length === 0
            ? 'No saved requests yet — click Save to create one.'
            : `${list.length} saved · click to load, hover for rename/delete`}
        </Caption1>
        {list.length === 0 ? (
          <div className={s.libraryEmpty}>
            <BookmarkMultiple20Regular style={{ width: 32, height: 32, opacity: 0.4 }} />
            <div style={{ marginTop: 8 }}>Nothing here yet.</div>
          </div>
        ) : (
          <div className={s.libraryList}>
            {list.map(entry => {
              const isCurrent = entry.id === currentId;
              const isRenaming = entry.id === renameId;
              return (
                <div
                  key={entry.id}
                  className={mergeClasses(s.libraryRow)}
                  style={isCurrent ? { borderColor: tokens.colorBrandStroke1, backgroundColor: tokens.colorBrandBackground2 } : undefined}
                  onClick={() => {
                    if (isRenaming) return;
                    onLoad(entry);
                    setOpen(false);
                  }}
                >
                  <Badge
                    appearance="tint"
                    color={isCurrent ? 'brand' : 'informative'}
                    size="small"
                    style={{ fontFamily: tokens.fontFamilyMonospace, minWidth: 80, justifyContent: 'center' }}
                  >
                    {modeIdShort(entry.modeId)}
                  </Badge>
                  <div style={{ minWidth: 0 }}>
                    {isRenaming ? (
                      <Input
                        value={renameDraft}
                        onChange={(_, d) => setRenameDraft(d.value)}
                        size="small"
                        style={{ width: '100%' }}
                        onClick={(e) => e.stopPropagation()}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') { e.preventDefault(); onCommitRename(); }
                          if (e.key === 'Escape') { e.preventDefault(); setRenameId(null); }
                        }}
                        onBlur={onCommitRename}
                        autoFocus
                      />
                    ) : (
                      <>
                        <div className={s.libraryRowName}>
                          {isCurrent && (
                            <Checkmark16Filled
                              style={{ color: tokens.colorBrandForeground1, marginRight: 4 }}
                            />
                          )}
                          {entry.name}
                        </div>
                        <div className={s.libraryRowMeta}>
                          {tableNameFromState(entry.state)} · saved {timeAgo(entry.savedAt)}
                        </div>
                      </>
                    )}
                  </div>
                  <div style={{ display: 'flex', gap: 2, alignItems: 'center' }}>
                    <Tooltip content="Rename" relationship="label">
                      <Button
                        appearance="subtle"
                        size="small"
                        icon={<Edit20Regular />}
                        onClick={(e) => {
                          e.stopPropagation();
                          setRenameId(entry.id);
                          setRenameDraft(entry.name);
                        }}
                        aria-label={`Rename ${entry.name}`}
                      />
                    </Tooltip>
                    <Tooltip content="Delete" relationship="label">
                      <Button
                        appearance="subtle"
                        size="small"
                        icon={<Delete20Regular />}
                        onClick={(e) => {
                          e.stopPropagation();
                          if (confirm(`Delete "${entry.name}"?`)) remove(entry.id);
                        }}
                        aria-label={`Delete ${entry.name}`}
                      />
                    </Tooltip>
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </PopoverSurface>
    </Popover>
  );
}

// ── helpers ──────────────────────────────────────────────────────────

/** Short chip label for the mode id — drops the leading "retrieve-" / "execute-"
 *  prefix for compact display, and falls back to the raw id for unknown modes. */
function modeIdShort(modeId: string): string {
  return modeId.replace(/^(retrieve|execute|manage)-/, '');
}

function timeAgo(epochMs: number): string {
  const sec = Math.round((Date.now() - epochMs) / 1000);
  if (sec < 60) return 'just now';
  if (sec < 3600) return `${Math.floor(sec / 60)} min ago`;
  if (sec < 86400) return `${Math.floor(sec / 3600)} h ago`;
  if (sec < 86400 * 7) return `${Math.floor(sec / 86400)} d ago`;
  return new Date(epochMs).toISOString().slice(0, 10);
}

// Use the same icon name as the file's primary, for callers that want
// to import individually rather than spreading via the button.
export { Save20Regular as SaveIcon, Open16Regular as OpenIcon };
