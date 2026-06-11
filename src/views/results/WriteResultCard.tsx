// WriteResultCard — terse success/failure card for write-side operations.
//
// Until now the Results tab routed every "single" mode through the
// RecordDetailCard, which works great for Retrieve Single (a record came
// back in the body) but yields a near-empty pane for write modes that
// return 204 No Content (Delete, Update without return=rep, Merge,
// Upsert "updated" path). The user gets no narrative confirmation of
// what just happened.
//
// This card replaces / wraps that for the write group:
//
//   • A short, operation-specific success line ("Updated <name>",
//     "Merged <sub> into <tgt>", etc.)
//   • Status pill + ms + bytes (same vocabulary as the grid footer)
//   • Action chips: Copy ID / Copy URL / Open in new tab
//   • If the response carries a body (Create 201 with return=rep,
//     Update 200, etc.) we include the existing RecordDetailCard below
//     so the user can drill into the freshly-saved row.
//
// Designed to render correctly for 200 / 201 / 204 success and also the
// common 412 / 404 / 409 failure cases — same component, different copy.

import { Caption1, Badge, tokens, Button, Tooltip, mergeClasses } from '@fluentui/react-components';
import {
  CheckmarkCircle20Filled,
  ErrorCircle20Filled,
  Warning20Filled,
  Copy20Regular,
  Open20Regular,
  BranchFork20Filled,
} from '@fluentui/react-icons';
import { useStudioStyles } from '../../primitives/styles';
import { StatusPill } from '../../primitives/StatusPill';
import { RecordDetailCard } from '../detail/RecordDetailCard';
import { ENV } from '../../mock/environment';
import { findTable } from '../../mock/metadata';
import type { ExecResult } from '../../engine/dataverseExecutor';

export type WriteOperation =
  | 'create'
  | 'update'
  | 'upsert'
  | 'delete'
  | 'merge'
  | 'associate'
  | 'disassociate'
  | 'action'
  | 'function'
  | 'workflow'
  | 'manage-file'
  | 'manage-image'
  | 'manage-attachment';

export interface WriteResultContext {
  operation: WriteOperation;
  /** Target table logical name — drives RecordDetailCard if body present. */
  table: string;
  /** GUID of the record operated on. For Create this comes from the
   *  response's `OData-EntityId` header (the executor sets it). For all
   *  other ops, the mode passes its own recordId. */
  recordId?: string | null;
  /** Resolved primary name when known — drives the success narrative. */
  recordName?: string | null;
  /** Delete single-property: which column was cleared. */
  clearedColumn?: string | null;
  /** Merge: the two record names for the success line. */
  targetName?: string | null;
  subordinateName?: string | null;
  /**
   * Associate / Disassociate — the nav property name (e.g.
   * `contact_customer_accounts`) and the count of targets that were
   * (dis)associated. `targetNames` is best-effort: only names captured
   * via the live RecordPicker land here.
   */
  navProperty?: string;
  targetCount?: number;
  targetNames?: string[];
  /**
   * Execute-group context — action / function / workflow.
   *   • operationName: the CSDL action name (e.g. 'WinOpportunity') or the
   *     workflow's display name.
   *   • boundEntity: the entity the action/function is bound to (when bound).
   */
  operationName?: string;
  boundEntity?: string;
  /**
   * Binary-group context — manage-file / manage-image.
   *   • fileOperation: upload / download / delete
   *   • columnName: the file/image column the operation targeted
   *   • fileName: optional payload filename (upload only)
   */
  fileOperation?: 'upload' | 'download' | 'delete';
  columnName?: string;
  fileName?: string;
}

export interface WriteResultCardProps {
  result: ExecResult;
  ctx: WriteResultContext;
}

export function WriteResultCard({ result, ctx }: WriteResultCardProps) {
  const s = useStudioStyles();
  const tbl = findTable(ctx.table);
  const entityDisplay = tbl?.displayName ?? ctx.table;

  // Resolve the affected record's GUID. Three fallbacks in priority order:
  //   1. `ctx.recordId` — the caller knows it (Update / Delete / Merge /
  //      Upsert-with-GUID-key all do).
  //   2. `OData-EntityId` response header — set by the host when a fresh
  //      record was created; the URL contains the new GUID in `(<guid>)`.
  //   3. The response body itself — for Create, PPTB's `dataverseAPI.create`
  //      resolves with `{ id: '<guid>', ... }`. We also check for the
  //      entity-specific PK field (`accountid`, `contactid`, etc.) in case
  //      the body carries return=representation. This is critical because
  //      PPTB doesn't surface `OData-EntityId` through its headers map for
  //      `create()`, so without this fallback Create would have an empty
  //      affordance row.
  const odataEntityId = result.headers['OData-EntityId'] ?? result.headers['odata-entityid'];
  const resolvedId =
    ctx.recordId ??
    extractGuidFromOdataEntityId(odataEntityId) ??
    extractGuidFromBody(result.body, tbl?.primaryKey);

  // Status family — drives icon + tone.
  const family: 'success' | 'warning' | 'danger' =
    result.status >= 200 && result.status < 300
      ? 'success'
      : result.status === 412 || result.status === 404 || result.status === 409
        ? 'warning'
        : 'danger';
  const Icon =
    family === 'success'
      ? CheckmarkCircle20Filled
      : family === 'warning'
        ? Warning20Filled
        : ErrorCircle20Filled;
  const tone =
    family === 'success'
      ? tokens.colorPaletteGreenForeground1
      : family === 'warning'
        ? tokens.colorPaletteDarkOrangeForeground1
        : tokens.colorPaletteRedForeground1;

  const narrative = buildNarrative(ctx, result, resolvedId, entityDisplay);
  const recordUrl =
    resolvedId && tbl
      ? `https://${ENV.host}/api/data/v9.2/${tbl.entitySetName}(${resolvedId})`
      : null;

  // The body — if Create/Update/Upsert returned a representation, we
  // render the recursive RecordDetailCard underneath the success header.
  // For 204 / null body / non-record body, we just show the header.
  const body = result.body;
  const hasReturnedRecord =
    !!body &&
    typeof body === 'object' &&
    !Array.isArray(body) &&
    Object.keys(body as Record<string, unknown>).length > 0 && // Cheap heuristic: real records carry @odata.etag OR the table's primary key
    ('@odata.etag' in (body as object) || (tbl && tbl.primaryKey in (body as object)));

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
      {/* Success / failure header card */}
      <div
        className={mergeClasses(s.inlineCard)}
        style={{
          padding: 16,
          background:
            family === 'success'
              ? tokens.colorPaletteGreenBackground1
              : family === 'warning'
                ? tokens.colorPaletteDarkOrangeBackground1
                : tokens.colorPaletteRedBackground1,
          border:
            family === 'success'
              ? `1px solid ${tokens.colorPaletteGreenBorderActive}`
              : family === 'warning'
                ? `1px solid ${tokens.colorPaletteDarkOrangeBorderActive}`
                : `1px solid ${tokens.colorPaletteRedBorderActive}`,
        }}
      >
        <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
          <Icon style={{ width: 28, height: 28, color: tone, flexShrink: 0 }} />
          <div style={{ flex: 1, minWidth: 0 }}>
            <div style={{ fontSize: 15, fontWeight: 600, marginBottom: 2 }}>{narrative.title}</div>
            {narrative.body && (
              <Caption1 style={{ color: tokens.colorNeutralForeground2 }}>
                {narrative.body}
              </Caption1>
            )}
          </div>
          <span style={{ display: 'inline-flex', alignItems: 'center', gap: 8, flexShrink: 0 }}>
            <StatusPill
              status={
                family === 'success' ? 'success' : family === 'warning' ? 'warning' : 'danger'
              }
              code={result.status}
              ms={result.ms}
            />
            <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
              {(result.bytes / 1024).toFixed(1)} KB
            </Caption1>
          </span>
        </div>

        {/* Affordances row.
            Delete's affordances are narrower: the record is gone, so the
            URL would 404 and Open-in-tab is a footgun. We keep Copy ID
            for audit/log purposes only. */}
        {resolvedId && (
          <div
            style={{
              display: 'flex',
              gap: 6,
              marginTop: 12,
              flexWrap: 'wrap',
              paddingTop: 12,
              borderTop: `1px solid ${tokens.colorNeutralStroke3}`,
            }}
          >
            <code
              style={{
                fontFamily: tokens.fontFamilyMonospace,
                fontSize: 11,
                color: tokens.colorNeutralForeground3,
                padding: '4px 8px',
                background: tokens.colorNeutralBackground1,
                borderRadius: tokens.borderRadiusSmall,
                alignSelf: 'center',
              }}
            >
              {tbl?.primaryKey ?? 'id'}: {resolvedId}
            </code>
            <Tooltip content="Copy record GUID" relationship="label">
              <Button
                size="small"
                appearance="subtle"
                icon={<Copy20Regular />}
                onClick={() => navigator.clipboard?.writeText(resolvedId)}
              >
                Copy ID
              </Button>
            </Tooltip>
            {/* URL / Open suppressed for Delete success — the URL 404s
                immediately after delete. For single-property "clear" the
                row survives, so we keep them. */}
            {recordUrl && !(ctx.operation === 'delete' && !ctx.clearedColumn) && (
              <>
                <Tooltip content="Copy record URL" relationship="label">
                  <Button
                    size="small"
                    appearance="subtle"
                    icon={<Copy20Regular />}
                    onClick={() => navigator.clipboard?.writeText(recordUrl)}
                  >
                    Copy URL
                  </Button>
                </Tooltip>
                <Tooltip content="Open the record's Web API URL in a new tab" relationship="label">
                  <Button
                    size="small"
                    appearance="subtle"
                    icon={<Open20Regular />}
                    onClick={() => window.open(recordUrl, '_blank', 'noopener')}
                  >
                    Open
                  </Button>
                </Tooltip>
              </>
            )}
          </div>
        )}
      </div>

      {/* If the server returned a representation, drill the full record card. */}
      {hasReturnedRecord && (
        <div>
          <Caption1
            style={{
              display: 'block',
              marginBottom: 6,
              color: tokens.colorNeutralForeground3,
              fontWeight: 600,
              textTransform: 'uppercase',
              letterSpacing: '0.04em',
              fontSize: 10,
            }}
          >
            Returned representation
          </Caption1>
          <RecordDetailCard
            record={body as Record<string, unknown>}
            entityLogical={ctx.table}
            level={0}
          />
        </div>
      )}
    </div>
  );
}

// ──────────────────────────────────────────────────────────────
// Narrative composition — short, accurate sentence per op + status
// ──────────────────────────────────────────────────────────────
function buildNarrative(
  ctx: WriteResultContext,
  result: ExecResult,
  resolvedId: string | null,
  entityDisplay: string,
): { title: string; body: string | null } {
  const name = ctx.recordName?.trim() || null;
  const ok = result.ok;

  // ── Failure framings — common error codes get specific copy ──
  if (!ok) {
    if (result.status === 412) {
      const why =
        ctx.operation === 'upsert'
          ? "precondition failed — the record's existence state didn't match your selection"
          : ctx.operation === 'delete' || ctx.operation === 'update'
            ? 'precondition failed — the ETag is stale (someone else changed the record)'
            : 'precondition failed';
      return { title: 'Operation rejected', body: why };
    }
    if (result.status === 404) {
      return {
        title: 'Record not found',
        body: `No ${entityDisplay} record matched ${resolvedId ?? 'the supplied id'}.`,
      };
    }
    if (result.status === 409) {
      return {
        title: 'Conflict',
        body: 'A duplicate-detection rule or related conflict blocked the operation.',
      };
    }
    if (result.status === 401) {
      return {
        title: 'Not authenticated',
        body: 'Your session lost the bearer token. Reconnect and retry.',
      };
    }
    if (result.status === 403) {
      return { title: 'Forbidden', body: 'The calling user lacks the required privilege.' };
    }
    if (result.status >= 500) {
      return {
        title: 'Server error',
        body: 'Dataverse returned a 5xx. Inspect the JSON body for details.',
      };
    }
    return {
      title: `Failed · ${result.status} ${result.statusText}`,
      body: 'See the JSON / Headers tabs for details.',
    };
  }

  // ── Success framings — operation-specific ──
  const who = name ? (
    <>{name}</>
  ) : resolvedId ? (
    `record ${resolvedId.slice(0, 8)}…`
  ) : (
    `${entityDisplay} record`
  );
  const whoStr = typeof who === 'string' ? who : (name as string);

  switch (ctx.operation) {
    case 'create':
      return {
        title: `Created ${entityDisplay}`,
        body: resolvedId
          ? `New record ${whoStr ? `"${whoStr}"` : ''} created with id ${resolvedId.slice(0, 8)}…`
          : 'Record created.',
      };
    case 'update':
      return {
        title:
          result.status === 200
            ? `Updated ${entityDisplay} (return=representation)`
            : `Updated ${entityDisplay}`,
        body: name
          ? `Saved changes to "${name}".`
          : resolvedId
            ? `Saved changes to ${resolvedId.slice(0, 8)}…`
            : null,
      };
    case 'upsert':
      // Per the Web API docs: server returns 201 if it created, 204 if it
      // updated. BUT — PPTB's `dataverseAPI.update()` resolves with `void`
      // and our executor hardcodes 204, so we can't actually distinguish
      // the two paths from inside DRS. Don't pretend we can.
      //
      // When status is unambiguous (e.g. 201, or 200 with body), we still
      // use the precise wording; otherwise hedge with "completed".
      if (result.status === 201) {
        return {
          title: `Upsert: created new ${entityDisplay}`,
          body: name
            ? `New record "${name}".`
            : resolvedId
              ? `id ${resolvedId.slice(0, 8)}…`
              : null,
        };
      }
      return {
        title: `Upsert completed on ${entityDisplay}`,
        body: name
          ? `Server resolved the request against "${name}". The PPTB host doesn't surface the distinguishing 201 (created) vs 204 (updated) — copy the URL + body from the Code tab and run outside PPTB if you need to tell which path ran.`
          : resolvedId
            ? `Server resolved the request against id ${resolvedId.slice(0, 8)}…`
            : 'Server resolved the request.',
      };
    case 'delete':
      if (ctx.clearedColumn) {
        return {
          title: `Cleared column on ${entityDisplay}`,
          body: `Set ${ctx.clearedColumn} to null on ${name ? `"${name}"` : `id ${resolvedId?.slice(0, 8) ?? '?'}…`}.`,
        };
      }
      return {
        title: `Deleted ${entityDisplay}`,
        body: name
          ? `"${name}" has been removed. Related records cascaded per CascadeConfiguration.`
          : 'Record removed.',
      };
    case 'merge':
      return {
        title: `Merged ${entityDisplay}`,
        body:
          ctx.subordinateName && ctx.targetName
            ? `"${ctx.subordinateName}" merged into "${ctx.targetName}". Subordinate deactivated; related rows re-parented.`
            : 'Merge completed. Subordinate deactivated; related rows re-parented.',
      };
    case 'associate': {
      const n = ctx.targetCount ?? 0;
      const navLabel = ctx.navProperty ? `via ${ctx.navProperty}` : '';
      const sourceFrag = name
        ? `"${name}"`
        : resolvedId
          ? `id ${resolvedId.slice(0, 8)}…`
          : `the ${entityDisplay} record`;
      // List up to 3 target names inline; rest become "and N more".
      const namesShown = (ctx.targetNames ?? []).slice(0, 3);
      const remaining = Math.max(0, n - namesShown.length);
      const targetsFrag =
        namesShown.length > 0
          ? `${namesShown.map((t) => `"${t}"`).join(', ')}${remaining > 0 ? ` and ${remaining} more` : ''}`
          : n > 0
            ? `${n} target${n === 1 ? '' : 's'}`
            : 'the target';
      return {
        title:
          n > 1 ? `Associated ${n} records to ${entityDisplay}` : `Associated to ${entityDisplay}`,
        body: `Linked ${targetsFrag} to ${sourceFrag} ${navLabel}.`.replace(/\s+\./, '.'),
      };
    }
    case 'disassociate': {
      const n = ctx.targetCount ?? 0;
      const navLabel = ctx.navProperty ? `via ${ctx.navProperty}` : '';
      const sourceFrag = name
        ? `"${name}"`
        : resolvedId
          ? `id ${resolvedId.slice(0, 8)}…`
          : `the ${entityDisplay} record`;
      return {
        title:
          n > 1
            ? `Disassociated ${n} records from ${entityDisplay}`
            : `Disassociated from ${entityDisplay}`,
        body:
          n > 0
            ? `Removed ${n} link${n === 1 ? '' : 's'} from ${sourceFrag} ${navLabel}.`.replace(
                /\s+\./,
                '.',
              )
            : `Cleared the lookup on ${sourceFrag} ${navLabel}.`.replace(/\s+\./, '.'),
      };
    }
    case 'action': {
      const opName = ctx.operationName ?? 'action';
      const boundFrag = ctx.boundEntity
        ? ` on ${name ? `${ctx.boundEntity} "${name}"` : `the ${ctx.boundEntity} record`}`
        : '';
      return {
        title: `Action: ${opName}`,
        body:
          result.status === 200 || result.status === 201
            ? `Executed ${opName}${boundFrag}. Server returned a result body — see JSON / Results tabs.`
            : `Executed ${opName}${boundFrag}. Server returned 204 No Content — no return value.`,
      };
    }
    case 'function': {
      const opName = ctx.operationName ?? 'function';
      const boundFrag = ctx.boundEntity
        ? ` bound to ${name ? `${ctx.boundEntity} "${name}"` : `a ${ctx.boundEntity} record`}`
        : '';
      return {
        title: `Function: ${opName}`,
        body: `Invoked ${opName}${boundFrag}. Result body in JSON / Results tabs.`,
      };
    }
    case 'workflow': {
      const opName = ctx.operationName ?? 'workflow';
      const targetFrag = name
        ? ` on "${name}"`
        : resolvedId
          ? ` on record ${resolvedId.slice(0, 8)}…`
          : '';
      return {
        title: `Workflow: ${opName}`,
        body: `Triggered "${opName}"${targetFrag}. Workflow steps run asynchronously — check the System Jobs grid for progress.`,
      };
    }
    case 'manage-file':
    case 'manage-image':
    case 'manage-attachment': {
      // Binary operations always return 501 inside PPTB (the host doesn't
      // expose the file/image/attachment upload + chunked-download APIs). The
      // failure case is the EXPECTED path here — we frame the narrative as
      // "request was built; run it externally" rather than "something went
      // wrong".
      const kind =
        ctx.operation === 'manage-image'
          ? 'image'
          : ctx.operation === 'manage-attachment'
            ? ctx.columnName === 'documentbody'
              ? 'annotation'
              : 'attachment'
            : 'file';
      const op = ctx.fileOperation ?? 'upload';
      const colFrag = ctx.columnName ? ` · column ${ctx.columnName}` : '';
      const targetFrag = name
        ? ` on ${entityDisplay} "${name}"`
        : resolvedId
          ? ` on record ${resolvedId.slice(0, 8)}…`
          : '';
      const kindLabel = kind.charAt(0).toUpperCase() + kind.slice(1);
      // 501 Not Implemented from the executor's hostNotSupported() is the
      // documented "use the Code tab" path; we frame it positively.
      if (result.status === 501) {
        return {
          title: `${kindLabel} ${op} — request built`,
          body:
            `PPTB's dataverseAPI doesn't expose the ${kind} ${op} endpoint. ` +
            `Your request is ready to run externally — copy a snippet from the Code tab ` +
            `(fetch / curl / C# / PowerShell) and execute it from there.`,
        };
      }
      return {
        title: `${kindLabel} ${op}${colFrag}`,
        body: result.ok
          ? `${kindLabel} ${op} completed${targetFrag}.`
          : `${op} request failed${targetFrag}. See the JSON / Headers tabs for details.`,
      };
    }
  }
}

/** Pulls the GUID out of an `OData-EntityId` header value:
 *    "https://<host>/api/data/v9.2/accounts(<guid>)" → "<guid>" */
function extractGuidFromOdataEntityId(header: string | undefined | null): string | null {
  if (!header) return null;
  const m = header.match(/\(([0-9a-fA-F-]{36})\)/);
  return m ? m[1] : null;
}

/** Extracts the new record's GUID from the response body.
 *
 * PPTB's `dataverseAPI.create` resolves with `{ id: '<guid>' }` even on a
 * minimal 201 (no return=representation). With return=representation on,
 * the body is the full record with the entity-typed primary key field
 * (e.g. `accountid`, `contactid`). We check both forms — `id` first
 * (matches PPTB's resolved shape), then the table's actual primaryKey
 * column.
 *
 * Returns null if the body isn't a record-shaped object. */
function extractGuidFromBody(body: unknown, primaryKey?: string): string | null {
  if (!body || typeof body !== 'object' || Array.isArray(body)) return null;
  const obj = body as Record<string, unknown>;
  const candidate =
    (typeof obj.id === 'string' && /^[0-9a-fA-F-]{36}$/.test(obj.id) ? obj.id : null) ??
    (primaryKey &&
    typeof obj[primaryKey] === 'string' &&
    /^[0-9a-fA-F-]{36}$/.test(obj[primaryKey] as string)
      ? (obj[primaryKey] as string)
      : null);
  return candidate ?? null;
}

// Suppress unused imports kept for parity
void BranchFork20Filled;
