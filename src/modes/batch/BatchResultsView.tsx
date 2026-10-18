// Per-request results for the $batch composer (#9), plus an inspector for a
// raw multipart response pasted from another client.

import { useState } from 'react';
import {
  Button,
  Caption1,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  Textarea,
  tokens,
} from '@fluentui/react-components';
import { ClipboardPaste20Regular } from '@fluentui/react-icons';
import { StatusPill, type StatusKind } from '../../primitives/StatusPill';
import { MethodPill } from '../../primitives/MethodPill';
import type { BatchRunResult, ItemOutcome } from '../../engine/batch/batchExecutor';
import { parseBatchResponse, type ParsedBatchResponse } from '../../engine/batch/multipart';
import type { BatchMethod } from '../../engine/batch/batchModel';

export interface ResultRow {
  key: string;
  contentId: string;
  method?: BatchMethod;
  path?: string;
  label?: string;
  outcome: ItemOutcome;
  status?: number;
  statusText?: string;
  headers?: Record<string, string>;
  body?: unknown;
}

const mono = { fontFamily: tokens.fontFamilyMonospace, fontSize: 12 };

function statusKind(r: ResultRow): StatusKind {
  if (r.outcome !== 'done') return r.outcome === 'rolled-back' ? 'warning' : 'info';
  if (r.status === undefined) return 'info';
  return r.status >= 400 ? 'danger' : 'success';
}

function errorMessage(body: unknown): string | null {
  if (body && typeof body === 'object' && 'error' in body) {
    const e = (body as { error?: { message?: string; code?: string } }).error;
    if (e?.message) return e.code ? `${e.code}: ${e.message}` : e.message;
  }
  return null;
}

function ResultRowView({ r }: { r: ResultRow }) {
  const err = errorMessage(r.body);
  const entityId = r.headers?.['odata-entityid'] ?? r.headers?.['OData-EntityId'];
  const outcomeText =
    r.outcome === 'rolled-back'
      ? 'Rolled back'
      : r.outcome === 'not-run'
        ? 'Not run'
        : r.outcome === 'unknown'
          ? 'No result'
          : null;
  return (
    <details
      style={{
        borderBottom: `1px solid ${tokens.colorNeutralStroke3}`,
        padding: '6px 4px',
      }}
    >
      <summary
        style={{
          display: 'flex',
          gap: 8,
          alignItems: 'center',
          cursor: 'pointer',
          listStyle: 'none',
        }}
      >
        <code style={{ width: 40, color: tokens.colorNeutralForeground3 }}>${r.contentId}</code>
        {r.method && <MethodPill method={r.method} size="sm" />}
        <code
          style={{
            flex: 1,
            minWidth: 0,
            overflow: 'hidden',
            textOverflow: 'ellipsis',
            whiteSpace: 'nowrap',
          }}
        >
          {r.label ? `${r.label} — ` : ''}
          {r.path ?? ''}
        </code>
        {outcomeText && r.status === undefined ? (
          <StatusPill status={statusKind(r)} code={outcomeText} />
        ) : (
          <StatusPill
            status={statusKind(r)}
            code={`${r.status ?? '—'}${r.statusText && (r.status ?? 0) >= 300 ? ` ${r.statusText}` : ''}`}
          />
        )}
        {outcomeText && r.status !== undefined && <Caption1>{outcomeText}</Caption1>}
      </summary>
      <div style={{ padding: '8px 0 4px 48px' }}>
        {err && (
          <MessageBar intent="error" layout="multiline" style={{ marginBottom: 8 }}>
            <MessageBarBody>{err}</MessageBarBody>
          </MessageBar>
        )}
        {entityId && (
          <div style={{ ...mono, marginBottom: 6 }}>
            OData-EntityId: <code>{entityId}</code>
          </div>
        )}
        {r.headers && Object.keys(r.headers).length > 0 && (
          <pre style={{ ...mono, margin: '0 0 8px', whiteSpace: 'pre-wrap' }}>
            {Object.entries(r.headers)
              .map(([k, v]) => `${k}: ${v}`)
              .join('\n')}
          </pre>
        )}
        {r.body !== undefined && (
          <pre
            style={{
              ...mono,
              margin: 0,
              whiteSpace: 'pre-wrap',
              wordBreak: 'break-all',
              background: tokens.colorNeutralBackground1,
              border: `1px solid ${tokens.colorNeutralStroke2}`,
              borderRadius: 4,
              padding: 8,
              maxHeight: 360,
              overflow: 'auto',
            }}
          >
            {typeof r.body === 'string' ? r.body : JSON.stringify(r.body, null, 2)}
          </pre>
        )}
        {r.body === undefined && r.outcome === 'done' && (
          <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>No response body.</Caption1>
        )}
        {r.outcome === 'rolled-back' && (
          <Caption1 block style={{ color: tokens.colorNeutralForeground3 }}>
            The change set failed, so Dataverse undid this request with the rest of it.
          </Caption1>
        )}
        {r.outcome === 'unknown' && (
          <Caption1 block style={{ color: tokens.colorNeutralForeground3 }}>
            The host call failed as a whole (see the message above), so there’s no result for this
            request.
          </Caption1>
        )}
        {r.outcome === 'not-run' && (
          <Caption1 block style={{ color: tokens.colorNeutralForeground3 }}>
            An earlier failure stopped the batch before this request ran.
          </Caption1>
        )}
      </div>
    </details>
  );
}

export function BatchResultsView({
  run,
  rows,
  extraTitle,
}: {
  run: BatchRunResult | null;
  rows: ResultRow[];
  extraTitle?: string;
}) {
  const [pasted, setPasted] = useState('');
  const [inspected, setInspected] = useState<ParsedBatchResponse[] | null>(null);
  const [parseError, setParseError] = useState<string | null>(null);

  const inspect = () => {
    try {
      setInspected(parseBatchResponse(pasted));
      setParseError(null);
    } catch (e) {
      setInspected(null);
      setParseError(e instanceof Error ? e.message : String(e));
    }
  };

  const counts = rows.reduce(
    (acc, r) => {
      if (r.outcome === 'done' && r.status !== undefined && r.status < 400) acc.ok += 1;
      else if (r.outcome === 'done') acc.failed += 1;
      else acc.other += 1;
      return acc;
    },
    { ok: 0, failed: 0, other: 0 },
  );

  return (
    <div style={{ maxWidth: 1000 }}>
      {run ? (
        <>
          <div
            style={{
              display: 'flex',
              gap: 10,
              alignItems: 'center',
              marginBottom: 12,
              flexWrap: 'wrap',
            }}
          >
            <StatusPill
              status={run.ok ? 'success' : 'danger'}
              code={run.ok ? 200 : 'Failed'}
              ms={run.ms}
            />
            <span style={{ fontSize: 13 }}>
              {counts.ok} succeeded · {counts.failed} failed
              {counts.other ? ` · ${counts.other} rolled back or not run` : ''}
              {run.segments.length > 1 ? ` · ${run.segments.length} host calls` : ''}
            </span>
            {extraTitle && <Caption1>{extraTitle}</Caption1>}
          </div>
          {run.segments.map((s, i) =>
            s.error || s.changesetError ? (
              <MessageBar key={i} intent="error" layout="multiline" style={{ marginBottom: 8 }}>
                <MessageBarBody>
                  <MessageBarTitle>
                    {s.transaction ? 'Change set' : 'Requests'} ${s.contentIds[0]}
                    {s.contentIds.length > 1
                      ? `–$${s.contentIds[s.contentIds.length - 1]}`
                      : ''}{' '}
                    failed
                  </MessageBarTitle>
                  {s.error ??
                    `${s.changesetError!.status} ${s.changesetError!.statusText ?? ''} — ${
                      errorMessage(s.changesetError!.body) ?? 'see the response body below'
                    }`}
                  {s.changesetError?.body !== undefined && !errorMessage(s.changesetError.body) && (
                    <pre style={{ ...mono, whiteSpace: 'pre-wrap' }}>
                      {JSON.stringify(s.changesetError.body, null, 2)}
                    </pre>
                  )}
                </MessageBarBody>
              </MessageBar>
            ) : null,
          )}
          <div
            style={{
              background: tokens.colorNeutralBackground1,
              border: `1px solid ${tokens.colorNeutralStroke2}`,
              borderRadius: 6,
              padding: '0 8px',
            }}
          >
            {rows.map((r) => (
              <ResultRowView key={r.key} r={r} />
            ))}
          </div>
        </>
      ) : (
        <Caption1 block style={{ color: tokens.colorNeutralForeground3, marginBottom: 16 }}>
          Send the batch to see each request’s status, headers and body here.
        </Caption1>
      )}

      <details style={{ marginTop: 24 }}>
        <summary style={{ cursor: 'pointer', fontSize: 13 }}>
          <ClipboardPaste20Regular style={{ verticalAlign: 'middle', marginRight: 6 }} />
          Inspect a raw $batch response from another client
        </summary>
        <div style={{ marginTop: 8 }}>
          <Textarea
            value={pasted}
            onChange={(_, d) => setPasted(d.value)}
            placeholder="--batchresponse_…"
            rows={8}
            resize="vertical"
            textarea={{ style: mono, spellCheck: false }}
            style={{ width: '100%' }}
            aria-label="Raw $batch response"
          />
          <Button style={{ marginTop: 8 }} onClick={inspect} disabled={!pasted.trim()}>
            Parse response
          </Button>
          {parseError && (
            <MessageBar intent="error" style={{ marginTop: 8 }}>
              <MessageBarBody>{parseError}</MessageBarBody>
            </MessageBar>
          )}
          {inspected && (
            <div style={{ marginTop: 8 }}>
              {inspected.map((r, i) => (
                <ResultRowView
                  key={i}
                  r={{
                    key: String(i),
                    contentId: r.contentId ?? String(i + 1),
                    label: r.changeset !== undefined ? `change set ${r.changeset + 1}` : undefined,
                    outcome: 'done',
                    status: r.status,
                    statusText: r.statusText,
                    headers: r.headers,
                    body: r.body,
                  }}
                />
              ))}
            </div>
          )}
        </div>
      </details>
    </div>
  );
}
