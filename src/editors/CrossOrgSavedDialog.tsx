// Saved requests on other connections (#6).
//
// Saved requests are stored per Dataverse connection
// (`drs:saved-requests:v1:<host>`), so switching orgs shows a different
// library. This dialog lists the libraries of the other connections this
// browser profile has used and copies or moves picked entries into the
// current one: debug a request in Dev, then open UAT and bring it over.
//
// Copies keep their id, so copying the same entry again updates it instead
// of duplicating it. A name already used here (by another entry) gets a
// " (2)" suffix, since names are unique within a library (#125).

import { useEffect, useMemo, useState } from 'react';
import {
  Badge,
  Button,
  Caption1,
  Checkbox,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Dropdown,
  Field,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  Option,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import { Dismiss20Regular, Globe20Regular } from '@fluentui/react-icons';
import {
  MAX_SAVED,
  getOrgScope,
  importedEntryProblem,
  listOtherOrgLibraries,
  removeFromOrgLibrary,
  tableNameFromState,
  type OrgLibrary,
  type SavedRequest,
} from '../state/savedRequests';
import { useSavedRequests } from '../state/useSavedRequests';

const useStyles = makeStyles({
  list: {
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
    maxHeight: '360px',
    overflowY: 'auto',
    marginTop: '8px',
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    paddingTop: '4px',
  },
  row: {
    display: 'grid',
    gridTemplateColumns: 'auto auto 1fr',
    gap: '8px',
    alignItems: 'center',
    padding: '2px 4px',
    borderRadius: tokens.borderRadiusMedium,
    ':hover': { backgroundColor: tokens.colorNeutralBackground1Hover },
  },
  name: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  meta: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    fontFamily: tokens.fontFamilyMonospace,
  },
  empty: {
    padding: '32px 16px',
    textAlign: 'center',
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
});

interface Status {
  intent: 'success' | 'warning' | 'error';
  title: string;
  text: string;
}

export interface CrossOrgSavedDialogProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
}

export function CrossOrgSavedDialog({ open, onOpenChange }: CrossOrgSavedDialogProps) {
  const s = useStyles();
  const { saved, importEntries } = useSavedRequests();
  const current = getOrgScope();
  // Re-read the other buckets each time the dialog opens and after a move.
  const [libraries, setLibraries] = useState<OrgLibrary[]>([]);
  const [scope, setScope] = useState<string>('');
  const [picked, setPicked] = useState<Set<string>>(new Set());
  const [status, setStatus] = useState<Status | null>(null);

  const refresh = (keepScope?: string) => {
    const libs = listOtherOrgLibraries();
    setLibraries(libs);
    setScope(
      keepScope && libs.some((l) => l.scope === keepScope) ? keepScope : (libs[0]?.scope ?? ''),
    );
    setPicked(new Set());
  };

  useEffect(() => {
    if (!open) return;
    refresh();
    setStatus(null);
  }, [open]);

  const library = libraries.find((l) => l.scope === scope);
  const entries = useMemo(
    () => [...(library?.entries ?? [])].sort((a, b) => b.savedAt - a.savedAt),
    [library],
  );
  const allPicked = entries.length > 0 && entries.every((e) => picked.has(e.id));

  const toggle = (id: string, on: boolean) =>
    setPicked((p) => {
      const next = new Set(p);
      if (on) next.add(id);
      else next.delete(id);
      return next;
    });

  const bringHere = (move: boolean) => {
    if (!library) return;
    // Entries whose ids or names would change the URL stay where they are,
    // as on import (#221).
    const blocked: string[] = [];
    const chosen = entries.filter((e) => {
      if (!picked.has(e.id)) return false;
      const problem = importedEntryProblem(e);
      if (problem) blocked.push(`"${e.name}": ${problem}`);
      return !problem;
    });
    if (!chosen.length) {
      if (blocked.length) {
        setStatus({ intent: 'error', title: "Couldn't copy", text: blocked.join(' ') });
      }
      return;
    }
    const here = new Set(saved.map((e) => e.id));
    const newCount = chosen.filter((e) => !here.has(e.id)).length;
    if (saved.length + newCount > MAX_SAVED) {
      setStatus({
        intent: 'error',
        title: 'Not enough room',
        text: `This connection has ${MAX_SAVED - saved.length} free slot(s) of ${MAX_SAVED} and ${newCount} of the picked request(s) are new here. Pick fewer or delete some saved requests first.`,
      });
      return;
    }
    // Names are unique per library: rename a copy whose name is taken here
    // by a different entry.
    const taken = new Map(saved.map((e) => [e.name.toLowerCase(), e.id]));
    let renamed = 0;
    const incoming: SavedRequest[] = chosen.map((e) => {
      const owner = taken.get(e.name.toLowerCase());
      let name = e.name;
      if (owner && owner !== e.id) {
        name = freeName(e.name, (n) => taken.has(n.toLowerCase()));
        renamed++;
      }
      taken.set(name.toLowerCase(), e.id);
      return name === e.name ? e : { ...e, name };
    });
    const r = importEntries(incoming);
    if (!r.ok) {
      setStatus({ intent: 'error', title: "Couldn't copy", text: r.error ?? '' });
      return;
    }
    const notes = [`${r.added} added, ${r.updated} updated, ${r.unchanged} already up to date.`];
    if (renamed) notes.push(`${renamed} renamed with a (n) suffix because the name was in use.`);
    if (blocked.length) notes.push(`Not ${move ? 'moved' : 'copied'}: ${blocked.join(' ')}`);
    let intent: Status['intent'] = renamed || blocked.length ? 'warning' : 'success';
    if (move) {
      try {
        removeFromOrgLibrary(library.scope, new Set(chosen.map((e) => e.id)));
        notes.push(`Removed from ${library.scope}.`);
      } catch (e) {
        intent = 'warning';
        notes.push(
          `Couldn't remove them from ${library.scope}: ${e instanceof Error ? e.message : String(e)}`,
        );
      }
      refresh(library.scope);
    } else {
      setPicked(new Set());
    }
    setStatus({ intent, title: move ? 'Moved' : 'Copied', text: notes.join(' ') });
  };

  return (
    <Dialog open={open} onOpenChange={(_, d) => onOpenChange(d.open)}>
      <DialogSurface style={{ maxWidth: 620 }}>
        <DialogBody>
          <DialogTitle
            action={
              <Button
                appearance="subtle"
                icon={<Dismiss20Regular />}
                aria-label="close"
                onClick={() => onOpenChange(false)}
              />
            }
          >
            Saved requests on other connections
          </DialogTitle>
          <DialogContent>
            <Caption1 style={{ display: 'block', color: tokens.colorNeutralForeground2 }}>
              Each Dataverse connection has its own saved requests. Copy or move requests saved
              under another connection into{' '}
              <code style={{ fontFamily: tokens.fontFamilyMonospace }}>{current}</code>. They are
              copied as-is: the tables, columns and record ids they use must exist here.
            </Caption1>
            {libraries.length === 0 ? (
              <div className={s.empty}>
                <Globe20Regular style={{ width: 32, height: 32, opacity: 0.4 }} />
                <div style={{ marginTop: 8 }}>
                  No other connection has saved requests in this browser profile.
                </div>
              </div>
            ) : (
              <>
                <Field label="Connection" style={{ marginTop: 12 }}>
                  <Dropdown
                    value={library ? `${library.scope} (${library.entries.length})` : ''}
                    selectedOptions={scope ? [scope] : []}
                    onOptionSelect={(_, d) => {
                      setScope(d.optionValue ?? '');
                      setPicked(new Set());
                      setStatus(null);
                    }}
                  >
                    {libraries.map((l) => (
                      <Option
                        key={l.scope}
                        value={l.scope}
                        text={`${l.scope} (${l.entries.length})`}
                      >
                        {`${l.scope} (${l.entries.length})`}
                      </Option>
                    ))}
                  </Dropdown>
                </Field>
                <div className={s.list}>
                  <Checkbox
                    label={`${picked.size} of ${entries.length} selected`}
                    checked={allPicked ? true : picked.size > 0 ? 'mixed' : false}
                    onChange={(_, d) =>
                      setPicked(d.checked === true ? new Set(entries.map((e) => e.id)) : new Set())
                    }
                  />
                  {entries.map((e) => (
                    <div key={e.id} className={s.row}>
                      <Checkbox
                        checked={picked.has(e.id)}
                        onChange={(_, d) => toggle(e.id, d.checked === true)}
                        aria-label={`Select ${e.name}`}
                      />
                      <Badge
                        appearance="tint"
                        color="informative"
                        size="small"
                        style={{
                          fontFamily: tokens.fontFamilyMonospace,
                          minWidth: 80,
                          justifyContent: 'center',
                        }}
                      >
                        {e.modeId.replace(/^(retrieve|execute|manage)-/, '')}
                      </Badge>
                      <div style={{ minWidth: 0 }}>
                        <div className={s.name}>{e.name}</div>
                        <div className={s.meta}>
                          {tableNameFromState(e.state)} · saved{' '}
                          {new Date(e.savedAt).toISOString().slice(0, 10)}
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              </>
            )}
            {status && (
              <MessageBar layout="multiline" intent={status.intent} style={{ marginTop: 12 }}>
                <MessageBarBody>
                  <MessageBarTitle>{status.title}</MessageBarTitle>
                  {status.text}
                </MessageBarBody>
              </MessageBar>
            )}
          </DialogContent>
          <DialogActions>
            <Button appearance="subtle" onClick={() => onOpenChange(false)}>
              Close
            </Button>
            <Button disabled={picked.size === 0} onClick={() => bringHere(true)}>
              Move here
            </Button>
            <Button
              appearance="primary"
              disabled={picked.size === 0}
              onClick={() => bringHere(false)}
            >
              Copy here
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}

/** First of `name (2)`, `name (3)`, … not yet taken. */
function freeName(name: string, isTaken: (n: string) => boolean): string {
  const base = name.replace(/\s*\(\d+\)$/, '');
  for (let n = 2; ; n++) {
    const candidate = `${base} (${n})`;
    if (!isTaken(candidate)) return candidate;
  }
}
