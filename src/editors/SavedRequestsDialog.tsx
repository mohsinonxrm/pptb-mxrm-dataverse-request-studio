// Saved-requests UX: two surfaces in one file because they share state.
//
//   <SaveButton>         — dirty-aware button in the URL bar. Opens the
//                           save dialog with an auto-suggested name.
//   <SaveDialog>         — name input + Save / Cancel. If the name
//                           collides with an existing entry, shows an
//                           overwrite confirm.
//   <SavedLibraryButton> — opens a popover listing all saved entries
//                           with click-to-load + rename + delete, and
//                           Export / Import of the whole library (#105).
//
// Library lives in localStorage (see savedRequests.ts). Per-mode load
// routing is the caller's responsibility — the dialog just hands back
// the SavedRequest; the mode decides what to do with `state`.

import { useEffect, useMemo, useRef, useState } from 'react';
import { Globe20Regular } from '@fluentui/react-icons';
import { CrossOrgSavedDialog } from './CrossOrgSavedDialog';
import {
  Dialog,
  DialogTrigger,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogActions,
  DialogContent,
  Popover,
  PopoverTrigger,
  PopoverSurface,
  Button,
  Input,
  Caption1,
  Badge,
  tokens,
  Tooltip,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  MessageBarActions,
  Switch,
  Menu,
  MenuTrigger,
  MenuButton,
  MenuPopover,
  MenuList,
  MenuItem,
  mergeClasses,
  makeStyles,
} from '@fluentui/react-components';
import {
  Save20Regular,
  Save20Filled,
  BookmarkMultiple20Regular,
  Delete20Regular,
  Edit20Regular,
  Dismiss20Regular,
  Open16Regular,
  Checkmark16Filled,
  ArrowDownload16Regular,
  ArrowUpload16Regular,
  DocumentArrowRight16Regular,
} from '@fluentui/react-icons';
import { isEmbedded, saveFileSafe, tbHost } from '../host/pptbBridge';
import {
  buildCollection,
  collectionFileName,
  countRequests,
  COLLECTION_FORMAT_LABELS,
  type CollectionFormat,
} from '../engine/collectionFormats';
import { currentBaseUrl, loadTablesFor, savedToCollectionItem } from '../engine/savedToCollection';
import {
  autoSuggestName,
  buildSavedExport,
  parseSavedImport,
  MAX_SAVED,
  newSavedId,
  getOrgScope,
  sameSaveTarget,
  tableNameFromState,
  type SavedRequest,
  type SavedModeId,
} from '../state/savedRequests';
import { useSavedRequests } from '../state/useSavedRequests';
import { findRequestType } from '../registry/requestTypes';
import { useConfirm } from '../shell/dialogs';

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
  transfer: {
    display: 'flex',
    flexDirection: 'column',
    gap: '6px',
    marginTop: '6px',
    paddingTop: '6px',
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    maxWidth: '420px',
  },
  transferActions: {
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
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
  // Library at MAX_SAVED and this save would add an entry: ask which one
  // to drop instead of evicting one silently (#100).
  const [confirmReplace, setConfirmReplace] = useState(false);
  // The typed name belongs to a saved request of ANOTHER mode (names are
  // unique across modes, the library is per mode). We suggest a new name
  // instead of overwriting it; overwriting stays an explicit choice (#125).
  const [crossMode, setCrossMode] = useState<{ existing: SavedRequest; typed: string } | null>(
    null,
  );
  const { save, capacity, findById, findByName, error } = useSavedRequests();

  // The remembered entry only counts while it targets the same table /
  // action as the current draft. After a target switch it belongs to a
  // different request: don't pre-fill its name or overwrite it silently.
  // A same-name save then goes through the overwrite confirm below (#58).
  const lastSaved = lastSavedId ? findById(lastSavedId) : undefined;
  const ownSavedId = lastSaved && sameSaveTarget(lastSaved.state, state) ? lastSaved.id : undefined;

  // Auto-suggest name whenever the dialog opens. If the user previously
  // saved this request, default to the existing name so a quick re-save
  // is one click + Enter.
  useEffect(() => {
    if (!open) return;
    if (ownSavedId) {
      const existing = findById(ownSavedId);
      setName(existing?.name ?? autoSuggestName(modeId, state));
    } else {
      setName(autoSuggestName(modeId, state));
    }
    setConfirmOverwrite(false);
    setCrossMode(null);
  }, [open, state, modeId, ownSavedId, findById]);

  const doSave = (overwriteId?: string, replaceId?: string, nameOverride?: string) => {
    const trimmed = (nameOverride ?? name).trim();
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
    const r = save(entry, replaceId ? { replaceId } : undefined);
    if (!r.ok) {
      if (r.full) setConfirmReplace(true);
      return; // other errors surface via useSavedRequests().error
    }
    onSaved(entry);
    setOpen(false);
  };

  const onClickSave = () => {
    const trimmed = name.trim();
    if (!trimmed) return;
    // Overwrite paths:
    //  • Same id (the entry we previously saved as): silent overwrite.
    //  • Same name, different id (collision): prompt confirm.
    if (ownSavedId && findById(ownSavedId)?.name === trimmed) {
      doSave(ownSavedId);
      return;
    }
    const byName = findByName(trimmed);
    if (byName && byName.modeId !== modeId) {
      // Collision with another mode's entry (not in this mode's Saved
      // list) → suggest a free name; overwriting is the explicit option.
      setCrossMode({ existing: byName, typed: trimmed });
      setName(freeName(trimmed, findByName));
      return;
    }
    if (byName) {
      // Collision with a DIFFERENT entry → confirm overwrite or rename.
      setConfirmOverwrite(true);
      return;
    }
    if (capacity.full) {
      setConfirmReplace(true);
      return;
    }
    doSave();
  };

  // Render: the trigger button itself (in the URL bar) + the dialog content.
  return (
    <Dialog
      open={open}
      onOpenChange={(_, d) => {
        setOpen(d.open);
        setConfirmReplace(false);
      }}
    >
      <DialogTrigger disableButtonEnhancement>
        <Tooltip
          content={dirty ? 'Save current request (unsaved changes)' : 'No changes since last save'}
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
            <Caption1
              style={{ display: 'block', marginBottom: 6, color: tokens.colorNeutralForeground2 }}
            >
              Saved requests live in this browser&apos;s storage, scoped to{' '}
              <code style={{ fontFamily: tokens.fontFamilyMonospace }}>{getOrgScope()}</code>. Other
              Dataverse connections have their own saved-request lists. Name auto-suggested from the
              current entity, select, and filter counts.
              {capacity.count >= capacity.max - 10 &&
                ` ${capacity.count} of ${capacity.max} saved requests used for this connection (all modes).`}
            </Caption1>
            <Input
              aria-label="Request name"
              value={name}
              onChange={(_, d) => {
                setName(d.value);
                setConfirmOverwrite(false);
                setConfirmReplace(false);
                setCrossMode(null);
              }}
              placeholder="Untitled request"
              style={{ width: '100%' }}
              autoFocus
              onKeyDown={(e) => {
                if (e.key === 'Enter' && !confirmOverwrite && !confirmReplace) {
                  e.preventDefault();
                  onClickSave();
                }
              }}
            />
            {confirmOverwrite && (
              <MessageBar layout="multiline" intent="warning" style={{ marginTop: 12 }}>
                <MessageBarBody>
                  <MessageBarTitle>Name in use</MessageBarTitle>A saved request in this mode already
                  uses <code>{name.trim()}</code>. Save again to overwrite it, or change the name
                  above and Save.
                </MessageBarBody>
              </MessageBar>
            )}
            {crossMode && (
              <MessageBar layout="multiline" intent="warning" style={{ marginTop: 12 }}>
                <MessageBarBody>
                  <MessageBarTitle>Name in use in another mode</MessageBarTitle>A saved{' '}
                  <strong>{modeName(crossMode.existing.modeId)}</strong> request already uses{' '}
                  <code>{crossMode.existing.name}</code> (names are shared across modes). The name
                  above was changed to a free one — Save to keep both.
                </MessageBarBody>
                <MessageBarActions>
                  <Button
                    size="small"
                    onClick={() => doSave(crossMode.existing.id, undefined, crossMode.typed)}
                  >
                    Overwrite the {modeName(crossMode.existing.modeId)} request
                  </Button>
                </MessageBarActions>
              </MessageBar>
            )}
            {confirmReplace && (
              <MessageBar layout="multiline" intent="warning" style={{ marginTop: 12 }}>
                <MessageBarBody>
                  <MessageBarTitle>Saved requests are full</MessageBarTitle>
                  This connection already has {capacity.max} saved requests (across all modes).
                  Delete some from the Saved list, save under an existing name to overwrite it, or
                  replace the oldest one
                  {capacity.oldest && (
                    <>
                      : <code>{capacity.oldest.name}</code> ({modeName(capacity.oldest.modeId)},
                      saved {timeAgo(capacity.oldest.savedAt)})
                    </>
                  )}
                  .
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
                if (confirmReplace) {
                  if (capacity.oldest) doSave(undefined, capacity.oldest.id);
                } else if (confirmOverwrite) {
                  const existing = findByName(name.trim());
                  if (existing) doSave(existing.id);
                } else {
                  onClickSave();
                }
              }}
              disabled={!name.trim() || (confirmReplace && !capacity.oldest)}
            >
              {confirmReplace ? 'Replace oldest' : confirmOverwrite ? 'Overwrite' : 'Save'}
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
  const askConfirm = useConfirm();
  // Keeps the popover open while the in-app delete confirm is up: clicks on
  // that dialog land outside the popover and would otherwise dismiss it.
  const confirmingRef = useRef(false);

  // Filter to this mode's entries AND the current org scope. The bucket
  // is already org-scoped via getSavedRequestsKey(), so this filter is a
  // belt-and-suspenders guard — if a corrupt entry slipped in via manual
  // migration or a future bulk-import feature, it stays hidden here.
  // Future modes get their own routing automatically: entries are tagged
  // with modeId at save time.
  const scope = getOrgScope();
  // "All modes" shows every mode's entries (names are unique across modes,
  // so a "Name in use" can point at one of them). Other modes' entries can
  // be renamed / deleted here but load only from their own mode (#125).
  const [allModes, setAllModes] = useState(false);
  // Other connections' libraries (#6).
  const [crossOrgOpen, setCrossOrgOpen] = useState(false);
  const modeCount = useMemo(
    () => saved.filter((e) => e.modeId === modeId && (!e.orgScope || e.orgScope === scope)).length,
    [saved, modeId, scope],
  );
  const list = useMemo(
    () =>
      saved
        .filter((e) => (allModes || e.modeId === modeId) && (!e.orgScope || e.orgScope === scope))
        .sort((a, b) => b.savedAt - a.savedAt),
    [saved, modeId, scope, allModes],
  );

  const onCommitRename = () => {
    if (!renameId) return;
    const trimmed = renameDraft.trim();
    if (trimmed) rename(renameId, trimmed);
    setRenameId(null);
  };

  return (
    <Popover
      open={open}
      onOpenChange={(_, d) => {
        if (!d.open && confirmingRef.current) return;
        setOpen(d.open);
        setRenameId(null);
      }}
    >
      <PopoverTrigger disableButtonEnhancement>
        <Tooltip content="Open saved requests" relationship="label">
          <Button
            appearance="subtle"
            icon={<BookmarkMultiple20Regular />}
            aria-label="Open saved requests"
          >
            Saved
            {modeCount > 0 && (
              <Badge appearance="tint" size="small" style={{ marginLeft: 6 }}>
                {modeCount}
              </Badge>
            )}
          </Button>
        </Tooltip>
      </PopoverTrigger>
      <PopoverSurface style={{ padding: 8 }}>
        {/* Rendered inside the surface: its portal counts as inside the popover. */}
        <CrossOrgSavedDialog open={crossOrgOpen} onOpenChange={setCrossOrgOpen} />
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
          <Caption1
            style={{ display: 'block', padding: '6px 10px', color: tokens.colorNeutralForeground3 }}
          >
            {list.length === 0
              ? 'No saved requests yet — click Save to create one.'
              : `${list.length} saved · click to load, hover for rename/delete`}
          </Caption1>
          <div style={{ display: 'flex', alignItems: 'center' }}>
            <Switch
              label="All modes"
              checked={allModes}
              onChange={(_, d) => setAllModes(d.checked)}
            />
            <Tooltip content="Saved requests on other connections" relationship="label">
              <Button
                appearance="subtle"
                size="small"
                icon={<Globe20Regular />}
                aria-label="Saved requests on other connections"
                onClick={() => setCrossOrgOpen(true)}
              />
            </Tooltip>
          </div>
        </div>
        {list.length === 0 ? (
          <div className={s.libraryEmpty}>
            <BookmarkMultiple20Regular style={{ width: 32, height: 32, opacity: 0.4 }} />
            <div style={{ marginTop: 8 }}>Nothing here yet.</div>
          </div>
        ) : (
          <div className={s.libraryList}>
            {list.map((entry) => {
              const isCurrent = entry.id === currentId;
              const isRenaming = entry.id === renameId;
              const otherMode = entry.modeId !== modeId;
              return (
                <div
                  key={entry.id}
                  className={mergeClasses(s.libraryRow)}
                  style={
                    isCurrent
                      ? {
                          borderColor: tokens.colorBrandStroke1,
                          backgroundColor: tokens.colorBrandBackground2,
                        }
                      : otherMode
                        ? { cursor: 'default' }
                        : undefined
                  }
                  onClick={() => {
                    if (isRenaming || otherMode) return;
                    onLoad(entry);
                    setOpen(false);
                  }}
                >
                  <Badge
                    appearance="tint"
                    color={isCurrent ? 'brand' : 'informative'}
                    size="small"
                    style={{
                      fontFamily: tokens.fontFamilyMonospace,
                      minWidth: 80,
                      justifyContent: 'center',
                    }}
                  >
                    {modeIdShort(entry.modeId)}
                  </Badge>
                  <div style={{ minWidth: 0 }}>
                    {isRenaming ? (
                      <Input
                        aria-label={`New name for ${entry.name}`}
                        value={renameDraft}
                        onChange={(_, d) => setRenameDraft(d.value)}
                        size="small"
                        style={{ width: '100%' }}
                        onClick={(e) => e.stopPropagation()}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') {
                            e.preventDefault();
                            onCommitRename();
                          }
                          if (e.key === 'Escape') {
                            e.preventDefault();
                            setRenameId(null);
                          }
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
                          {otherMode && ` · open ${modeName(entry.modeId)} to load`}
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
                          confirmingRef.current = true;
                          void askConfirm({
                            title: 'Delete saved request?',
                            message: `"${entry.name}" will be removed from your saved requests.`,
                            confirmLabel: 'Delete',
                            destructive: true,
                          }).then((ok) => {
                            confirmingRef.current = false;
                            if (ok) remove(entry.id);
                          });
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
        <SavedLibraryTransfer visible={list} allModes={allModes} modeId={modeId} />
      </PopoverSurface>
    </Popover>
  );
}

// ═════════════════════════════════════════════════════════════════════
//  Export / Import — footer of the Saved library popover (#105)
// ═════════════════════════════════════════════════════════════════════

const JSON_FILTER = [{ name: 'JSON', extensions: ['json'] }];

interface TransferStatus {
  intent: 'success' | 'warning' | 'error';
  title: string;
  text: string;
}

/**
 * Backs up / restores the whole library of the current connection (every
 * mode) as a JSON file. Uses the host's native dialogs inside PPTB and a
 * download / file input standalone. Results show inline — no native dialogs.
 */
function SavedLibraryTransfer({
  visible,
  allModes,
  modeId,
}: {
  /** Entries shown in the list (this mode, or every mode with All modes on). */
  visible: SavedRequest[];
  allModes: boolean;
  modeId: SavedModeId;
}) {
  const s = useStyles();
  const { saved, importEntries } = useSavedRequests();
  const [status, setStatus] = useState<TransferStatus | null>(null);
  const [busy, setBusy] = useState(false);
  const fileInput = useRef<HTMLInputElement>(null);

  const fail = (title: string, e: unknown) =>
    setStatus({ intent: 'error', title, text: e instanceof Error ? e.message : String(e) });

  const applyImport = (text: string) => {
    try {
      const { entries, invalid } = parseSavedImport(text);
      const r = importEntries(entries);
      if (!r.ok) {
        setStatus({ intent: 'error', title: "Couldn't import", text: r.error ?? '' });
        return;
      }
      const notes: string[] = [
        `${r.added} added, ${r.updated} updated, ${r.unchanged} already up to date.`,
      ];
      if (r.skippedFull > 0) {
        notes.push(
          `${r.skippedFull} not imported: this connection's library is full (${MAX_SAVED}).`,
        );
      }
      if (invalid > 0) {
        notes.push(`${invalid} item(s) in the file weren't valid saved requests and were skipped.`);
      }
      setStatus({
        intent: r.skippedFull > 0 || invalid > 0 ? 'warning' : 'success',
        title: 'Import finished',
        text: notes.join(' '),
      });
    } catch (e) {
      fail("Couldn't import", e);
    }
  };

  const onExport = async () => {
    setStatus(null);
    setBusy(true);
    try {
      const scope = getOrgScope();
      const fileName = `drs-saved-requests-${scope.replace(/[^a-z0-9.-]+/gi, '_')}-${new Date()
        .toISOString()
        .slice(0, 10)}.json`;
      const path = await saveFileSafe(fileName, buildSavedExport(saved, scope), JSON_FILTER);
      // Inside PPTB a null path means the user cancelled the save dialog.
      if (path || !isEmbedded()) {
        setStatus({
          intent: 'success',
          title: 'Exported',
          text: `${saved.length} saved request(s) written to ${path ?? fileName}.`,
        });
      }
    } catch (e) {
      fail("Couldn't export", e);
    } finally {
      setBusy(false);
    }
  };

  // Export the listed entries as a request collection for another HTTP
  // client (#27). Entries whose table isn't on this connection are skipped
  // and named in the result.
  const onExportCollection = async (fmt: CollectionFormat) => {
    setStatus(null);
    setBusy(true);
    try {
      await loadTablesFor(visible);
      const converted = visible.map(savedToCollectionItem);
      const items = converted.flatMap((c) => (c.ok ? [c.item] : []));
      const skipped = converted.flatMap((c) => (c.ok ? [] : [`${c.name} (${c.reason})`]));
      if (items.length === 0) {
        setStatus({
          intent: 'error',
          title: "Couldn't export",
          text: `None of the listed requests could be built. ${skipped.join('; ')}`,
        });
        return;
      }
      const scope = getOrgScope();
      const name = `Dataverse Request Studio · ${scope}${allModes ? '' : ` · ${modeIdShort(modeId)}`}`;
      const fileName = collectionFileName(
        fmt,
        `drs-${scope}${allModes ? '' : `-${modeIdShort(modeId)}`}`,
      );
      const text = buildCollection(fmt, { name, baseUrl: currentBaseUrl(), items });
      const path = await saveFileSafe(fileName, text, JSON_FILTER);
      if (path || !isEmbedded()) {
        const n = countRequests(items);
        setStatus({
          intent: skipped.length ? 'warning' : 'success',
          title: 'Exported',
          text:
            `${items.length} saved request(s) (${n} HTTP request(s)) written to ${path ?? fileName}. ` +
            'Set the accessToken variable to a bearer token before sending.' +
            (skipped.length ? ` Skipped: ${skipped.join('; ')}.` : ''),
        });
      }
    } catch (e) {
      fail("Couldn't export", e);
    } finally {
      setBusy(false);
    }
  };

  const onImport = async () => {
    setStatus(null);
    if (!isEmbedded() || !window.toolboxAPI) {
      fileInput.current?.click();
      return;
    }
    setBusy(true);
    try {
      const path = await tbHost.fileSystem.selectPath({
        type: 'file',
        title: 'Import saved requests',
        buttonLabel: 'Import',
        filters: JSON_FILTER,
      });
      if (path) applyImport(await tbHost.fileSystem.readText(path));
    } catch (e) {
      fail("Couldn't import", e);
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className={s.transfer}>
      <div className={s.transferActions}>
        <Tooltip
          content="Save every saved request for this connection (all modes) to a JSON file"
          relationship="description"
        >
          <Button
            appearance="subtle"
            size="small"
            icon={<ArrowDownload16Regular />}
            disabled={busy || saved.length === 0}
            onClick={() => void onExport()}
          >
            Export
          </Button>
        </Tooltip>
        <Menu>
          <MenuTrigger disableButtonEnhancement>
            <Tooltip
              content={`Export the ${visible.length} listed request(s) as a collection for another HTTP client`}
              relationship="description"
            >
              <MenuButton
                appearance="subtle"
                size="small"
                icon={<DocumentArrowRight16Regular />}
                disabled={busy || visible.length === 0}
              >
                Export as
              </MenuButton>
            </Tooltip>
          </MenuTrigger>
          <MenuPopover>
            <MenuList>
              {(Object.keys(COLLECTION_FORMAT_LABELS) as CollectionFormat[]).map((fmt) => (
                <MenuItem key={fmt} onClick={() => void onExportCollection(fmt)}>
                  {COLLECTION_FORMAT_LABELS[fmt]}
                </MenuItem>
              ))}
            </MenuList>
          </MenuPopover>
        </Menu>
        <Tooltip
          content="Add saved requests from an exported JSON file to this connection"
          relationship="description"
        >
          <Button
            appearance="subtle"
            size="small"
            icon={<ArrowUpload16Regular />}
            disabled={busy}
            onClick={() => void onImport()}
          >
            Import
          </Button>
        </Tooltip>
        <input
          ref={fileInput}
          type="file"
          accept=".json,application/json"
          hidden
          onChange={(e) => {
            const file = e.target.files?.[0];
            e.target.value = '';
            if (!file) return;
            file.text().then(applyImport, (err: unknown) => fail("Couldn't import", err));
          }}
        />
      </div>
      {status && (
        <MessageBar layout="multiline" intent={status.intent}>
          <MessageBarBody>
            <MessageBarTitle>{status.title}</MessageBarTitle>
            <span style={{ wordBreak: 'break-word' }}>{status.text}</span>
          </MessageBarBody>
        </MessageBar>
      )}
    </div>
  );
}

// ── helpers ──────────────────────────────────────────────────────────

/** Short chip label for the mode id — drops the leading "retrieve-" / "execute-"
 *  prefix for compact display, and falls back to the raw id for unknown modes. */
/** Full mode name for sentences ("Retrieve Multiple"); unknown ids fall back to the short id. */
function modeName(modeId: string): string {
  const t = findRequestType(modeId);
  return t.id === modeId ? t.name : modeIdShort(modeId);
}

function modeIdShort(modeId: string): string {
  return modeId.replace(/^(retrieve|execute|manage)-/, '');
}

/** First of `name (2)`, `name (3)`, … that no saved request uses (#125). */
function freeName(name: string, findByName: (n: string) => SavedRequest | undefined): string {
  const base = name.replace(/\s*\(\d+\)$/, '');
  for (let n = 2; ; n++) {
    const candidate = `${base} (${n})`;
    if (!findByName(candidate)) return candidate;
  }
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
