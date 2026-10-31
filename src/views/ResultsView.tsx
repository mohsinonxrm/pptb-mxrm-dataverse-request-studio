import { useState, type ReactNode } from 'react';
import {
  TabList,
  Tab,
  Button,
  Caption1,
  Text,
  tokens,
  Tooltip,
  Spinner,
  MessageBar,
  MessageBarBody,
  MessageBarActions,
} from '@fluentui/react-components';
import {
  Dismiss20Regular,
  Copy20Regular,
  ArrowDownload20Regular,
  DocumentSearch20Regular,
  ErrorCircle20Regular,
  LockClosed20Regular,
  ShieldDismiss20Regular,
  Filter20Regular,
  FilterDismiss20Regular,
  Link20Regular,
  DocumentCopy20Regular,
} from '@fluentui/react-icons';
import { useStudioStyles } from '../primitives/styles';
import { StatusPill } from '../primitives/StatusPill';
import { useAnnounceResult } from '../primitives/announce';
import {
  HEADER_CONSENT_DECLINED_CODE,
  HEADER_REJECTED_CODE,
  type ExecResult,
} from '../engine/dataverseExecutor';
import { pageSizeIsSent } from '../host/pptbBridge';
import { noStatusAdvice } from './results/noStatusAdvice';
import { ResultsGrid } from './results/ResultsGrid';
import { ResultsCommandBar } from './results/ResultsCommandBar';
import type { GridExportView } from './results/exportResults';
import { usePersistedSettings } from '../host/usePersistedSettings';
import { RecordDetailCard } from './detail/RecordDetailCard';
import { WriteResultCard, type WriteResultContext } from './results/WriteResultCard';
import type { ExpandSpec } from '../editors/ExpandEditor';
import type { OrderbySpec } from '../editors/OrderbyEditor';
import { copyText } from '../shell/copy';
import { ENV } from '../mock/environment';
import { useScopedEntities } from '../host/useScopedEntities';
import {
  describeMissingPrivilege,
  errorReport,
  missingPrivilege,
  notFoundSegment,
  suggestEntitySet,
} from './results/nextSteps';

/** The request URL as sent: modes pass it relative to the org host. */
function absoluteUrl(url: string | undefined): string | undefined {
  return url && url.startsWith('/') ? `https://${ENV.host}${url}` : url;
}

export type ResultsTab = 'data' | 'json' | 'headers';

export function ResultsView({
  result,
  table,
  mode,
  select,
  expand,
  orderby,
  requestUrl,
  isLoadingMore,
  onRefresh,
  onLoadMore,
  onRetrieveAll,
  onStopLoadingMore,
  loadMoreError,
  onRetryLoadMore,
  onDismissLoadMoreError,
  loadMoreNotice,
  onContinueLoadMore,
  onDismissLoadMoreNotice,
  writeContext,
  preferredColumnOrder,
  method = 'GET',
  onEditFilter,
  onRunWithoutFilter,
  onUseEntitySet,
  notice,
}: {
  result: ExecResult | null;
  table?: string;
  /** 'multi' shows DataGrid; 'single' shows a Detail card; 'next' shows Grid like multi */
  mode: 'multi' | 'single' | 'next';
  /** Active root $select — drives column order in the grid. */
  select?: string[];
  /** Active $expand — drives nested-column header labels in the grid. */
  expand?: ExpandSpec[];
  /** Active $orderby — drives sort-indicator arrows on the grid headers. */
  orderby?: OrderbySpec[];
  /** Full resolved URL for Copy URL in the command bar. */
  requestUrl?: string;
  /** True while a follow-up page is being fetched. */
  isLoadingMore?: boolean;
  /** Re-run the request. */
  onRefresh?: () => void;
  /** Fetch the next page via @odata.nextLink (called by the grid near bottom). */
  onLoadMore?: () => void;
  /** Walk every page until exhausted. */
  onRetrieveAll?: () => void;
  /** Stop an in-flight next-page fetch / Retrieve all, keeping rows loaded so far. */
  onStopLoadingMore?: () => void;
  /** Last next-page failure; rendered as an error bar above the grid. */
  loadMoreError?: string | null;
  /** Retry the failed next page. */
  onRetryLoadMore?: () => void;
  onDismissLoadMoreError?: () => void;
  /** Why Retrieve all paused with pages left (#223); rendered with Continue. */
  loadMoreNotice?: string | null;
  onContinueLoadMore?: () => void;
  onDismissLoadMoreNotice?: () => void;
  /**
   * Write-mode context. When supplied AND `mode === 'single'`, the
   * results pane renders WriteResultCard (operation-specific success /
   * failure narrative + actions) instead of the bare RecordDetailCard.
   * 204 No Content responses (Delete / Update / Merge / Upsert-updated)
   * still get a meaningful confirmation card this way.
   */
  writeContext?: WriteResultContext;
  /**
   * Column-order hint forwarded to ResultsGrid. Modes whose projection
   * isn't expressed via $select (Predefined Query reads from layoutxml)
   * pass this to keep the grid columns aligned with the source view's
   * column order. Ignored when `select` is non-empty.
   */
  preferredColumnOrder?: string[];
  /** HTTP method of the request, for the error report (#291). */
  method?: string;
  /** 0 rows: open the $filter node in the builder (#291). */
  onEditFilter?: () => void;
  /** 0 rows: drop the $filter and run again (#291). */
  onRunWithoutFilter?: () => void;
  /** 404 on an unknown entity set: switch the request to the suggested
   *  set (#291). Without it the suggestion is shown as text only. */
  onUseEntitySet?: (segment: string, entitySet: string) => void;
  /** A message bar above the results, e.g. "Ran without $filter". */
  notice?: ReactNode;
}) {
  const [tab, setTab] = useState<ResultsTab>('data');
  // Grid-local UI state. We hoist it here (not inside ResultsGrid) so the
  // command bar can drive it without prop-drilling through the grid.
  const [searchQuery, setSearchQuery] = useState('');
  // Row density is a persisted display setting (#16).
  const [settings, setSettings] = usePersistedSettings();
  const density = settings.gridDensity;
  // What the grid lists (columns + rows after search) — CSV export writes this.
  const [gridView, setGridView] = useState<GridExportView | null>(null);
  const [showFilterRow, setShowFilterRow] = useState(false);
  const s = useStudioStyles();
  useAnnounceResult(result);

  if (!result) {
    return (
      <div className={s.emptyState}>
        <DocumentSearch20Regular style={{ width: 48, height: 48, opacity: 0.5 }} />
        <Text size={400} weight="semibold">
          Nothing run yet
        </Text>
        <Caption1>
          Click <strong>Execute</strong> in the URL bar to run the request.
        </Caption1>
      </div>
    );
  }

  // Read-mode errors (multi / next) still go through the generic ErrorState —
  // it has the 401 / 403 / 404 icon families and a JSON-body preview. Write
  // modes (`writeContext` present) have operation-specific failure copy
  // baked into WriteResultCard (412 = "ETag stale" / "precondition failed",
  // 409 = duplicate-detect, etc.), so we fall through and let it render.
  if (!result.ok && !writeContext) {
    return (
      <ErrorState
        result={result}
        method={method}
        requestUrl={absoluteUrl(requestUrl)}
        onUseEntitySet={onUseEntitySet}
      />
    );
  }

  // 0 rows (#291): the grid's empty state offers the next step.
  const emptyActions = (
    <>
      {onEditFilter && (
        <Button icon={<Filter20Regular />} onClick={onEditFilter}>
          Edit $filter
        </Button>
      )}
      {onRunWithoutFilter && (
        <Button icon={<FilterDismiss20Regular />} onClick={onRunWithoutFilter}>
          Run without $filter
        </Button>
      )}
      {requestUrl && (
        <Button
          icon={<Link20Regular />}
          onClick={() => void copyText(absoluteUrl(requestUrl) ?? '', 'the request URL')}
        >
          Copy request URL
        </Button>
      )}
    </>
  );

  // Detect more-pages from the envelope's @odata.nextLink annotation,
  // and the server-emitted @odata.count (when $count=true was set).
  const env = result.body as {
    '@odata.nextLink'?: string;
    '@odata.count'?: number;
    value?: unknown[];
  } | null;
  const nextLink = env?.['@odata.nextLink'];
  const hasMore = !!nextLink;
  const serverTotal =
    typeof env?.['@odata.count'] === 'number' && env['@odata.count'] >= 0
      ? env['@odata.count']
      : undefined;

  const isGrid = tab === 'data' && mode !== 'single';

  if (result.outcome === 'empty') {
    return (
      <div
        style={{ display: 'flex', flexDirection: 'column', gap: 8, height: '100%', minHeight: 0 }}
      >
        <TabList
          selectedValue={tab}
          onTabSelect={(_, d) => setTab(d.value as ResultsTab)}
          size="small"
          appearance="subtle"
        >
          <Tab value="data">{mode === 'single' ? 'Detail' : 'Grid'}</Tab>
          <Tab value="json">JSON</Tab>
          <Tab value="headers">Headers</Tab>
        </TabList>
        <div className={s.emptyState}>
          <DocumentSearch20Regular style={{ width: 48, height: 48, opacity: 0.5 }} />
          <Text size={400} weight="semibold">
            No results
          </Text>
          <Caption1>The query returned 0 rows.</Caption1>
        </div>
      </div>
    );
  }

  return (
    // Fill exactly the mainBody content area. Earlier we used a
    // `calc(100vh - …)` value that didn't match the actual mainBody
    // height — when it was even a few pixels too tall, the outer
    // mainBody scrolled AND the grid scrolled, giving the user two
    // vertical scrollbars. `height: 100%` resolves to the flex-sized
    // mainBody, so there's only the grid's internal scroll.
    <div
      style={{
        display: 'flex',
        flexDirection: 'column',
        gap: 8,
        height: '100%',
        minHeight: 0,
      }}
    >
      {/* Tabs (status pill + ms + size live in the grid footer) */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
        <TabList
          selectedValue={tab}
          onTabSelect={(_, d) => setTab(d.value as ResultsTab)}
          size="small"
          appearance="subtle"
        >
          <Tab value="data">{mode === 'single' ? 'Detail' : 'Grid'}</Tab>
          <Tab value="json">JSON</Tab>
          <Tab value="headers">Headers</Tab>
        </TabList>
      </div>

      {isGrid && (
        <ResultsCommandBar
          hasMore={hasMore}
          isLoading={false}
          isLoadingMore={!!isLoadingMore}
          requestUrl={requestUrl}
          body={result.body}
          onRefresh={onRefresh ?? (() => {})}
          onRetrieveAll={onRetrieveAll}
          onStopLoadingMore={onStopLoadingMore}
          searchQuery={searchQuery}
          onSearchChange={setSearchQuery}
          showFilterRow={showFilterRow}
          onShowFilterRowChange={setShowFilterRow}
          density={density}
          onDensityChange={(d) => setSettings({ ...settings, gridDensity: d })}
          table={table}
          exportView={gridView}
        />
      )}

      {notice}

      {loadMoreError && (
        <MessageBar intent="error" layout="multiline">
          <MessageBarBody>{loadMoreError} Rows already loaded are kept.</MessageBarBody>
          <MessageBarActions
            containerAction={
              onDismissLoadMoreError && (
                <Button
                  appearance="transparent"
                  aria-label="Dismiss"
                  icon={<Dismiss20Regular />}
                  onClick={onDismissLoadMoreError}
                />
              )
            }
          >
            {onRetryLoadMore && <Button onClick={onRetryLoadMore}>Retry</Button>}
          </MessageBarActions>
        </MessageBar>
      )}

      {loadMoreNotice && (
        <MessageBar intent="info" layout="multiline">
          <MessageBarBody>{loadMoreNotice}</MessageBarBody>
          <MessageBarActions
            containerAction={
              onDismissLoadMoreNotice && (
                <Button
                  appearance="transparent"
                  aria-label="Dismiss"
                  icon={<Dismiss20Regular />}
                  onClick={onDismissLoadMoreNotice}
                />
              )
            }
          >
            {onContinueLoadMore && <Button onClick={onContinueLoadMore}>Continue</Button>}
          </MessageBarActions>
        </MessageBar>
      )}

      <div
        style={{
          flex: 1,
          minHeight: 0,
          display: 'flex',
          flexDirection: 'column',
          overflow: isGrid ? 'hidden' : 'auto',
        }}
      >
        {tab === 'data' &&
          (mode === 'single' ? (
            writeContext ? (
              <WriteResultCard result={result} ctx={writeContext} />
            ) : (
              <DetailPane body={result.body} table={table ?? ''} />
            )
          ) : (
            <ResultsGrid
              body={result.body}
              table={table}
              select={select}
              expand={expand}
              orderby={orderby}
              hasMore={hasMore}
              isLoadingMore={isLoadingMore}
              onLoadMore={onLoadMore}
              searchQuery={searchQuery}
              onClearSearch={() => setSearchQuery('')}
              showFilterRow={showFilterRow}
              density={density}
              status={{ code: result.status, ms: result.ms, bytes: result.bytes }}
              serverTotal={serverTotal}
              preferredColumnOrder={preferredColumnOrder}
              onViewChange={setGridView}
              emptyActions={emptyActions}
            />
          ))}
        {tab === 'json' && <JsonPane body={result.body} />}
        {tab === 'headers' && (
          <HeadersPane
            headers={result.headers}
            status={result.status}
            statusText={result.statusText}
          />
        )}
      </div>
    </div>
  );
}

// ResultsHeader was removed — status pill + ms + bytes + row count now live
// in the grid's bottom info bar (see ResultsGrid `infoBar`).

function DetailPane({ body, table }: { body: unknown; table: string }) {
  if (!body || typeof body !== 'object') return null;
  const rec = body as Record<string, unknown>;
  // The new RecordDetailCard handles:
  //   • scalar / lookup / formatted-value pairing in one grid
  //   • N:1 expanded nav objects as recursive sub-cards
  //   • 1:N / N:N expanded collections as a sub-grid (with per-row deep-
  //     expand toggles for chained expands like 1:N → N:1 → ...)
  //   • OData annotations (etag, context, etc.) tucked behind a collapsed
  //     section so the main grid stays clean
  //
  // No depth cap on recursion — depth is whatever the response carries.
  return <RecordDetailCard record={rec} entityLogical={table} level={0} />;
}

function JsonPane({ body }: { body: unknown }) {
  const json = JSON.stringify(body, null, 2);
  return (
    <div style={{ position: 'relative' }}>
      <Tooltip content="Copy JSON" relationship="label">
        <Button
          icon={<Copy20Regular />}
          appearance="subtle"
          size="small"
          style={{ position: 'absolute', top: 6, right: 6, zIndex: 1 }}
          onClick={() => void copyText(json, 'the JSON')}
        />
      </Tooltip>
      <pre
        style={{
          margin: 0,
          padding: 16,
          background: tokens.colorNeutralBackground1,
          border: `1px solid ${tokens.colorNeutralStroke2}`,
          borderRadius: tokens.borderRadiusMedium,
          fontFamily: tokens.fontFamilyMonospace,
          fontSize: 12,
          lineHeight: 1.6,
          color: tokens.colorNeutralForeground1,
          maxHeight: 540,
          overflow: 'auto',
        }}
      >
        {json}
      </pre>
    </div>
  );
}

function HeadersPane({
  headers,
  status,
  statusText,
}: {
  headers: Record<string, string>;
  status: number;
  statusText: string;
}) {
  const s = useStudioStyles();
  const lines = Object.entries(headers).filter(([, v]) => v !== '');
  const responseLine = status ? `HTTP/1.1 ${status} ${statusText}` : statusText;
  return (
    <div className={s.inlineCard} style={{ padding: 16 }}>
      <Caption1
        style={{
          display: 'block',
          fontFamily: tokens.fontFamilyMonospace,
          color: tokens.colorNeutralForeground3,
          marginBottom: 8,
        }}
      >
        {responseLine}
      </Caption1>
      {lines.length === 0 && (
        <Caption1 style={{ color: tokens.colorNeutralForeground3, fontStyle: 'italic' }}>
          Response headers aren&apos;t exposed by the PPTB host, so none can be shown here.
        </Caption1>
      )}
      <div
        style={{
          display: 'grid',
          gridTemplateColumns: '220px 1fr',
          gap: 6,
          fontFamily: tokens.fontFamilyMonospace,
          fontSize: 12,
        }}
      >
        {lines.map(([k, v]) => (
          <span key={k} style={{ display: 'contents' }}>
            <span style={{ color: tokens.colorBrandForeground1, fontWeight: 600 }}>{k}:</span>
            <span style={{ color: tokens.colorNeutralForeground1, wordBreak: 'break-all' }}>
              {v}
            </span>
          </span>
        ))}
      </div>
    </div>
  );
}

/**
 * True when the host refused the response because of its per-response size
 * cap ("... exceeds the maximum size", or the 16 MB Buffer limit) (#126).
 */
function isResponseSizeLimitError(message: string | undefined): boolean {
  return !!message && /16777216|exceeds the maximum size/i.test(message);
}

/** "Did you mean …?" for a 404 on an entity set the user mistyped (#291).
 *  Only tables inside the current scope are offered. Its own component, so
 *  the entity list is read only when there's a segment to match. */
function EntitySetSuggestion({
  segment,
  onUse,
}: {
  segment: string;
  onUse?: (segment: string, entitySet: string) => void;
}) {
  const { entities } = useScopedEntities();
  const match = suggestEntitySet(segment, entities);
  if (!match) return null;
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
      <Caption1 style={{ color: tokens.colorNeutralForeground1 }}>
        Did you mean <code>{match.entitySetName}</code> ({match.displayName || match.logicalName})?
      </Caption1>
      {onUse && (
        <Button size="small" onClick={() => onUse(segment, match.entitySetName)}>
          Use {match.entitySetName}
        </Button>
      )}
    </div>
  );
}

function ErrorState({
  result,
  method,
  requestUrl,
  onUseEntitySet,
}: {
  result: ExecResult;
  method?: string;
  requestUrl?: string;
  onUseEntitySet?: (segment: string, entitySet: string) => void;
}) {
  const Icon =
    result.status === 401
      ? LockClosed20Regular
      : result.status === 403
        ? ShieldDismiss20Regular
        : ErrorCircle20Regular;
  const err = (result.body as { error?: { code?: string; message?: string } } | null)?.error;
  // #126 — the host caps response size; say how to shrink the response.
  // Only point at the page size where Execute sends it (#219).
  const sizeHint = isResponseSizeLimitError(err?.message)
    ? 'The response is larger than the host allows. Narrow it: add $select on the table and ' +
      'inside each $expand, filter the expanded rows, or ' +
      (pageSizeIsSent()
        ? 'lower Prefer: odata.maxpagesize (or set $top when no page size is set).'
        : "set $top to return fewer rows (this ToolBox version can't send " +
          'Prefer: odata.maxpagesize; ToolBox 1.2.7+ can).')
    : null;
  // #126 — 0x80060888 (a generic OData code, so match on the message):
  // top-level $top with a collection $expand that has a nested $expand.
  const nestedExpandTopHint = /at top level while doing \$expand on nested one-to-many/i.test(
    err?.message ?? '',
  )
    ? 'Dataverse rejects top-level $top when a 1:N or N:N $expand has its own nested ' +
      '$expand. ' +
      (pageSizeIsSent()
        ? 'Remove $top and set Prefer: odata.maxpagesize instead to limit the rows per page.'
        : "Remove $top. This ToolBox version can't send Prefer: odata.maxpagesize, so pages " +
          "use Dataverse's 5,000-row default: narrow the rows with $filter, or use " +
          'ToolBox 1.2.7+ to set a page size.')
    : null;
  // Never sent: ToolBox refused the headers before calling Dataverse, so the
  // 403 / 400 isn't a privilege or request problem (#257).
  const notSent =
    err?.code === HEADER_CONSENT_DECLINED_CODE
      ? {
          title: 'Not sent — headers not approved',
          advice:
            "Power Platform ToolBox didn't send the request because its headers weren't approved, so nothing changed in Dataverse. Execute again and approve the headers when ToolBox asks.",
        }
      : err?.code === HEADER_REJECTED_CODE
        ? {
            title: 'Not sent — header rejected',
            advice:
              'Power Platform ToolBox rejected a request header. Fix it in the Headers or Prefer pane.',
          }
        : null;
  const title = notSent
    ? notSent.title
    : result.status === 401
      ? 'Unauthorized · 401'
      : result.status === 403
        ? 'Forbidden · 403'
        : result.status === 404
          ? 'Not Found · 404'
          : result.status === 501
            ? 'Not Implemented · 501'
            : result.status === 0
              ? 'Request failed'
              : `Error · ${result.status}`;
  // #291: name the missing privilege when the 403 message has one.
  const privilege = !notSent && result.status === 403 ? missingPrivilege(err?.message) : null;
  // #291: an entity set Dataverse couldn't resolve. Matched on the message,
  // since the host often reports no status for it.
  const segment = notSent ? null : notFoundSegment(err?.message);
  const advice = notSent
    ? notSent.advice
    : result.status === 401
      ? 'Your sign-in for this connection has expired or was revoked. Open Connection Manager in Power Platform ToolBox, sign in to the connection again, then run the request again.'
      : result.status === 403
        ? privilege
          ? describeMissingPrivilege(privilege)
          : 'Your user is missing a privilege this request needs. Check your security roles with a Dataverse administrator.'
        : result.status === 404
          ? 'Record (or entity) not found. Verify the GUID and entity set.'
          : result.status === 501
            ? 'The request was not sent. See the message below.'
            : result.status === 0
              ? noStatusAdvice(err)
              : 'Inspect the error message below for details.';
  // Any failure is 'danger', including a status the host didn't report (0).
  const kind: 'success' | 'danger' | 'warning' = !result.ok
    ? 'danger'
    : result.status >= 300
      ? 'warning'
      : 'success';
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
        <StatusPill status={kind} code={result.status || 'Error'} ms={result.ms} />
        <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
          {(result.bytes / 1024).toFixed(1)} KB
        </Caption1>
      </div>
      <div
        style={{
          display: 'flex',
          flexDirection: 'column',
          gap: 14,
          padding: 32,
          textAlign: 'center',
          alignItems: 'center',
          // Containment: `minWidth: 0` lets flex children shrink below their
          // intrinsic content size; `overflow: hidden` clamps any rogue
          // overflow (e.g. a long unbroken token like a GUID-with-trailing-
          // text); `maxWidth: 100%` keeps the card inside the parent column.
          minWidth: 0,
          maxWidth: '100%',
          overflow: 'hidden',
          border: `1px solid ${tokens.colorPaletteRedBorder2}`,
          borderRadius: tokens.borderRadiusLarge,
          background: tokens.colorPaletteRedBackground1,
        }}
      >
        <Icon style={{ width: 48, height: 48, color: tokens.colorPaletteRedForeground1 }} />
        <Text size={500} weight="semibold">
          {title}
        </Text>
        <Caption1 style={{ color: tokens.colorNeutralForeground2, maxWidth: 420 }}>
          {advice}
        </Caption1>
        {sizeHint && (
          <Caption1 style={{ color: tokens.colorNeutralForeground1, maxWidth: 420 }}>
            {sizeHint}
          </Caption1>
        )}
        {nestedExpandTopHint && (
          <Caption1 style={{ color: tokens.colorNeutralForeground1, maxWidth: 420 }}>
            {nestedExpandTopHint}
          </Caption1>
        )}
        {segment && <EntitySetSuggestion segment={segment} onUse={onUseEntitySet} />}
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', justifyContent: 'center' }}>
          <Button
            icon={<DocumentCopy20Regular />}
            onClick={() =>
              void copyText(errorReport({ result, method, requestUrl }), 'the error report')
            }
          >
            Copy error report
          </Button>
          {requestUrl && (
            <Button
              icon={<Link20Regular />}
              onClick={() => void copyText(requestUrl, 'the request URL')}
            >
              Copy request URL
            </Button>
          )}
        </div>
        {err && (
          // `whiteSpace: pre-wrap` keeps `\n` breaks but wraps the rest;
          // `wordBreak: break-word` + `overflowWrap: anywhere` break long
          // unbreakable strings (URLs, GUIDs, exception traces) so the
          // <pre> never expands past its container. `minWidth: 0` lets
          // it shrink inside the flex column.
          <pre
            style={{
              margin: 0,
              padding: 12,
              background: tokens.colorNeutralBackground1,
              border: `1px solid ${tokens.colorNeutralStroke2}`,
              borderRadius: tokens.borderRadiusMedium,
              fontFamily: tokens.fontFamilyMonospace,
              fontSize: 11,
              color: tokens.colorNeutralForeground1,
              textAlign: 'left',
              width: '100%',
              maxWidth: '100%',
              minWidth: 0,
              whiteSpace: 'pre-wrap',
              wordBreak: 'break-word',
              overflowWrap: 'anywhere',
              boxSizing: 'border-box',
            }}
          >
            {err.code ? `code: ${err.code}\nmessage: ${err.message}` : `message: ${err.message}`}
          </pre>
        )}
      </div>
    </div>
  );
}

// keep import warnings down
void Spinner;
void ArrowDownload20Regular;
