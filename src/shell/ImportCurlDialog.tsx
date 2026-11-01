// Import a cURL command into the matching mode (#271). Opened from the
// header's More actions menu, so it works from any mode.

import { useMemo, useRef, useState } from 'react';
import {
  Badge,
  Button,
  Caption1,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  Spinner,
  Textarea,
  tokens,
} from '@fluentui/react-components';
import { importCurl, type CurlImportResult } from '../engine/curlImport/curlImport';
import { openImport } from '../state/modeHandoff';
import { findRequestType } from '../registry/requestTypes';
import { useScopedEntities } from '../host/useScopedEntities';
import { metadata } from '../host/metadataProvider';
import { actions } from '../host/csdlProvider';
import { getEnv } from '../env/environment';

export function ImportCurlDialog({ open, onClose }: { open: boolean; onClose: () => void }) {
  const [text, setText] = useState('');
  const [result, setResult] = useState<CurlImportResult | null>(null);
  const [reading, setReading] = useState(false);
  // Bumped by every read, edit and close: a read that finishes after the text
  // changed or the dialog closed is dropped instead of shown for other text.
  const readSeq = useRef(0);
  // A command may name any table, in scope or not (#44).
  const { allEntities: entities } = useScopedEntities();
  const resolveEntitySet = useMemo(
    () => (set: string) => entities.find((e) => e.entitySetName === set)?.logicalName,
    [entities],
  );

  const close = () => {
    readSeq.current += 1;
    setReading(false);
    setText('');
    setResult(null);
    onClose();
  };

  const onRead = async () => {
    if (!text.trim()) return;
    const seq = ++readSeq.current;
    setReading(true);
    try {
      const r = await importCurl(text, {
        host: getEnv().host,
        resolveEntitySet,
        loadTable: (logical) => metadata.getTable(logical),
        findAction: (name) => actions.find(name),
      });
      if (seq === readSeq.current) setResult(r);
    } finally {
      if (seq === readSeq.current) setReading(false);
    }
  };

  const request = result?.errors.length ? undefined : result?.request;
  const modeName = request ? findRequestType(request.modeId).name : '';

  const onOpen = () => {
    if (!request) return;
    if (openImport(request)) close();
    else setResult((r) => r && { ...r, errors: ["Can't switch modes here. Open DRS in ToolBox."] });
  };

  return (
    <Dialog open={open} onOpenChange={(_, d) => !d.open && close()}>
      <DialogSurface style={{ maxWidth: 760 }}>
        <DialogBody>
          <DialogTitle>Import a cURL command</DialogTitle>
          <DialogContent style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
            <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
              Paste a <code>curl</code> command, for example from your browser&apos;s Copy as cURL
              or DRS&apos;s cURL code. It opens in the mode that sends it, filled in and not run.
              Authorization and cookies are left out: ToolBox signs you in.
            </Caption1>
            <Textarea
              aria-label="cURL command"
              value={text}
              onChange={(_, d) => {
                readSeq.current += 1;
                setReading(false);
                setText(d.value);
                setResult(null);
              }}
              rows={8}
              resize="vertical"
              placeholder={
                "curl 'https://contoso.crm.dynamics.com/api/data/v9.2/accounts?$top=5' \\\n  -H 'Accept: application/json'"
              }
              textarea={{ style: { fontFamily: tokens.fontFamilyMonospace, fontSize: 12 } }}
            />
            {reading && <Spinner size="tiny" label="Reading the command…" />}
            {result && result.errors.length > 0 && (
              <MessageBar intent="error" layout="multiline">
                <MessageBarBody>
                  <MessageBarTitle>Can&apos;t import this command</MessageBarTitle>
                  {result.errors.map((e, i) => (
                    <div key={i}>{e}</div>
                  ))}
                </MessageBarBody>
              </MessageBar>
            )}
            {request && (
              <div
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: 8,
                  padding: '8px 12px',
                  border: `1px solid ${tokens.colorNeutralStroke2}`,
                  borderRadius: tokens.borderRadiusMedium,
                  minWidth: 0,
                }}
              >
                <Badge appearance="tint" shape="rounded">
                  {result?.method}
                </Badge>
                <code
                  style={{
                    overflow: 'hidden',
                    textOverflow: 'ellipsis',
                    whiteSpace: 'nowrap',
                    minWidth: 0,
                    flex: 1,
                  }}
                  title={result?.path}
                >
                  {result?.path}
                </code>
                <Caption1 style={{ whiteSpace: 'nowrap' }}>
                  Opens in <strong>{modeName}</strong>
                </Caption1>
              </div>
            )}
            {result && result.warnings.length > 0 && (
              <MessageBar intent="warning" layout="multiline">
                <MessageBarBody>
                  {result.warnings.map((w, i) => (
                    <div key={i}>{w}</div>
                  ))}
                </MessageBarBody>
              </MessageBar>
            )}
            {result && result.notes.length > 0 && (
              <ul style={{ margin: 0, paddingLeft: 18, color: tokens.colorNeutralForeground3 }}>
                {result.notes.map((n, i) => (
                  <li key={i}>
                    <Caption1>{n}</Caption1>
                  </li>
                ))}
              </ul>
            )}
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={close}>
              Cancel
            </Button>
            {request ? (
              <Button appearance="primary" onClick={onOpen}>
                Open in {modeName}
              </Button>
            ) : (
              <Button
                appearance="primary"
                disabled={!text.trim() || reading}
                onClick={() => void onRead()}
              >
                Read command
              </Button>
            )}
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
