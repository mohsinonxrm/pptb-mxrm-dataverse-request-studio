// "Add from Saved" picker for the $batch composer (#9). Lists this
// connection's saved requests of every batchable mode; the chosen ones are
// rebuilt into batch requests by `savedToBatch`.

import { useMemo, useState } from 'react';
import {
  Button,
  Caption1,
  Checkbox,
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
import { useSavedRequests } from '../../state/useSavedRequests';
import { getOrgScope, type SavedRequest } from '../../state/savedRequests';
import { BATCHABLE_MODES } from '../../engine/batch/savedToBatch';
import { findRequestType } from '../../registry/requestTypes';

export function AddFromSavedDialog({
  open,
  onClose,
  onAdd,
}: {
  open: boolean;
  onClose: () => void;
  onAdd: (entries: SavedRequest[]) => void;
}) {
  const { saved } = useSavedRequests();
  const [picked, setPicked] = useState<Set<string>>(new Set());
  const [q, setQ] = useState('');
  const scope = getOrgScope();
  const list = useMemo(
    () =>
      saved
        .filter((e) => BATCHABLE_MODES.has(e.modeId) && (!e.orgScope || e.orgScope === scope))
        .filter((e) => !q.trim() || e.name.toLowerCase().includes(q.trim().toLowerCase()))
        .sort((a, b) => b.savedAt - a.savedAt),
    [saved, scope, q],
  );

  const close = () => {
    setPicked(new Set());
    setQ('');
    onClose();
  };

  return (
    <Dialog open={open} onOpenChange={(_, d) => !d.open && close()}>
      <DialogSurface style={{ maxWidth: 640 }}>
        <DialogBody>
          <DialogTitle>Add saved requests to the batch</DialogTitle>
          <DialogContent>
            <Input
              value={q}
              onChange={(_, d) => setQ(d.value)}
              placeholder="Filter by name"
              contentBefore={<Search20Regular />}
              style={{ width: '100%', marginBottom: 8 }}
              aria-label="Filter saved requests"
            />
            {list.length === 0 && (
              <Caption1 block style={{ color: tokens.colorNeutralForeground3, padding: '12px 0' }}>
                No saved requests on this connection that can go in a batch. Save a request from a
                Read, Write, Relate or Execute mode first.
              </Caption1>
            )}
            <div style={{ maxHeight: 360, overflowY: 'auto' }}>
              {list.map((e) => (
                <div key={e.id} style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                  <Checkbox
                    checked={picked.has(e.id)}
                    onChange={(_, d) =>
                      setPicked((p) => {
                        const next = new Set(p);
                        if (d.checked) next.add(e.id);
                        else next.delete(e.id);
                        return next;
                      })
                    }
                    label={
                      <span>
                        {e.name}{' '}
                        <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                          {findRequestType(e.modeId).name}
                        </Caption1>
                      </span>
                    }
                  />
                </div>
              ))}
            </div>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={close}>
              Cancel
            </Button>
            <Button
              appearance="primary"
              disabled={picked.size === 0}
              onClick={() => {
                // Keep the library's order (newest first) for what was ticked.
                onAdd(list.filter((e) => picked.has(e.id)));
                close();
              }}
            >
              Add {picked.size || ''}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
