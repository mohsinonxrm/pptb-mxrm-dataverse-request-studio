import { useState } from 'react';
import {
  TabList, Tab, Button, Caption1, Text, tokens, Tooltip, Spinner,
} from '@fluentui/react-components';
import {
  Copy20Regular, ArrowDownload20Regular, DocumentSearch20Regular,
  ErrorCircle20Regular, LockClosed20Regular, ShieldDismiss20Regular,
} from '@fluentui/react-icons';
import { useStudioStyles } from '../primitives/styles';
import { StatusPill } from '../primitives/StatusPill';
import type { ExecResult } from '../engine/dataverseExecutor';
import { ResultsGrid } from './results/ResultsGrid';
import { ResultsCommandBar } from './results/ResultsCommandBar';
import { RecordDetailCard } from './detail/RecordDetailCard';
import { WriteResultCard, type WriteResultContext } from './results/WriteResultCard';
import type { ExpandSpec } from '../editors/ExpandEditor';
import type { OrderbySpec } from '../editors/OrderbyEditor';

export type ResultsTab = 'data' | 'json' | 'headers';

export function ResultsView({
  result, table, mode, select, expand, orderby,
  requestUrl, isLoadingMore, onRefresh, onLoadMore, onRetrieveAll,
  writeContext, preferredColumnOrder,
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
}) {
  const [tab, setTab] = useState<ResultsTab>('data');
  // Grid-local UI state. We hoist it here (not inside ResultsGrid) so the
  // command bar can drive it without prop-drilling through the grid.
  const [searchQuery, setSearchQuery] = useState('');
  const [density, setDensity] = useState<'cozy' | 'compact'>('cozy');
  const s = useStudioStyles();

  if (!result) {
    return (
      <div className={s.emptyState}>
        <DocumentSearch20Regular style={{ width: 48, height: 48, opacity: 0.5 }} />
        <Text size={400} weight="semibold">Nothing run yet</Text>
        <Caption1>Click <strong>Execute</strong> in the URL bar to run the request.</Caption1>
      </div>
    );
  }

  // Read-mode errors (multi / next) still go through the generic ErrorState —
  // it has the 401 / 403 / 404 icon families and a JSON-body preview. Write
  // modes (`writeContext` present) have operation-specific failure copy
  // baked into WriteResultCard (412 = "ETag stale" / "precondition failed",
  // 409 = duplicate-detect, etc.), so we fall through and let it render.
  if (!result.ok && !writeContext) {
    return <ErrorState result={result} />;
  }

  // Detect more-pages from the envelope's @odata.nextLink annotation,
  // and the server-emitted @odata.count (when $count=true was set).
  const env = result.body as {
    '@odata.nextLink'?: string;
    '@odata.count'?: number;
    value?: unknown[];
  } | null;
  const nextLink = env?.['@odata.nextLink'];
  const hasMore = !!nextLink;
  const serverTotal = typeof env?.['@odata.count'] === 'number' && env['@odata.count'] >= 0
    ? env['@odata.count']
    : undefined;

  const isGrid = tab === 'data' && mode !== 'single';

  if (result.outcome === 'empty') {
    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 8, height: '100%', minHeight: 0 }}>
        <TabList selectedValue={tab} onTabSelect={(_, d) => setTab(d.value as ResultsTab)} size="small" appearance="subtle">
          <Tab value="data">{mode === 'single' ? 'Detail' : 'Grid'}</Tab>
          <Tab value="json">JSON</Tab>
          <Tab value="headers">Headers</Tab>
        </TabList>
        <div className={s.emptyState}>
          <DocumentSearch20Regular style={{ width: 48, height: 48, opacity: 0.5 }} />
          <Text size={400} weight="semibold">No results</Text>
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
    <div style={{
      display: 'flex',
      flexDirection: 'column',
      gap: 8,
      height: '100%',
      minHeight: 0,
    }}>
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
          searchQuery={searchQuery}
          onSearchChange={setSearchQuery}
          density={density}
          onDensityChange={setDensity}
        />
      )}

      <div style={{
        flex: 1, minHeight: 0, display: 'flex', flexDirection: 'column',
        overflow: isGrid ? 'hidden' : 'auto',
      }}>
        {tab === 'data' && (mode === 'single'
          ? (writeContext
              ? <WriteResultCard result={result} ctx={writeContext} />
              : <DetailPane body={result.body} table={table ?? ''} />)
          : <ResultsGrid
              body={result.body}
              table={table}
              select={select}
              expand={expand}
              orderby={orderby}
              hasMore={hasMore}
              isLoadingMore={isLoadingMore}
              onLoadMore={onLoadMore}
              searchQuery={searchQuery}
              density={density}
              status={{ code: result.status, ms: result.ms, bytes: result.bytes }}
              serverTotal={serverTotal}
              preferredColumnOrder={preferredColumnOrder}
            />)}
        {tab === 'json'   && <JsonPane body={result.body} />}
        {tab === 'headers' && <HeadersPane headers={result.headers} status={result.status} />}
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
  return (
    <RecordDetailCard
      record={rec}
      entityLogical={table}
      level={0}
    />
  );
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
          onClick={() => navigator.clipboard?.writeText(json)}
        />
      </Tooltip>
      <pre style={{
        margin: 0, padding: 16,
        background: tokens.colorNeutralBackground1,
        border: `1px solid ${tokens.colorNeutralStroke2}`,
        borderRadius: tokens.borderRadiusMedium,
        fontFamily: tokens.fontFamilyMonospace,
        fontSize: 12, lineHeight: 1.6, color: tokens.colorNeutralForeground1,
        maxHeight: 540, overflow: 'auto',
      }}>{json}</pre>
    </div>
  );
}

function HeadersPane({ headers, status }: { headers: Record<string, string>; status: number }) {
  const s = useStudioStyles();
  const lines = Object.entries(headers).filter(([, v]) => v !== '');
  const responseLine = `HTTP/1.1 ${status} ${status === 200 ? 'OK' : ''}`;
  return (
    <div className={s.inlineCard} style={{ padding: 16 }}>
      <Caption1 style={{ display: 'block', fontFamily: tokens.fontFamilyMonospace, color: tokens.colorNeutralForeground3, marginBottom: 8 }}>
        {responseLine}
      </Caption1>
      <div style={{ display: 'grid', gridTemplateColumns: '220px 1fr', gap: 6, fontFamily: tokens.fontFamilyMonospace, fontSize: 12 }}>
        {lines.map(([k, v]) => (
          <span key={k} style={{ display: 'contents' }}>
            <span style={{ color: tokens.colorBrandForeground1, fontWeight: 600 }}>{k}:</span>
            <span style={{ color: tokens.colorNeutralForeground1, wordBreak: 'break-all' }}>{v}</span>
          </span>
        ))}
      </div>
    </div>
  );
}

function ErrorState({ result }: { result: ExecResult }) {
  const Icon =
    result.status === 401 ? LockClosed20Regular :
    result.status === 403 ? ShieldDismiss20Regular :
    ErrorCircle20Regular;
  const err = (result.body as { error?: { code?: string; message?: string } } | null)?.error;
  const title =
    result.status === 401 ? 'Unauthorized · 401' :
    result.status === 403 ? 'Forbidden · 403' :
    result.status === 404 ? 'Not Found · 404' :
    `Error · ${result.status}`;
  const advice =
    result.status === 401 ? 'Sign in to refresh your access token, then re-run.' :
    result.status === 403 ? 'Caller is missing a required privilege. Check role assignments.' :
    result.status === 404 ? 'Record (or entity) not found. Verify the GUID and entity set.' :
    'Inspect the JSON body and headers tabs for details.';
  const kind: 'success' | 'danger' | 'warning' = result.status >= 400 ? 'danger' : result.status >= 300 ? 'warning' : 'success';
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
        <StatusPill status={kind} code={result.status} ms={result.ms} />
        <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
          {(result.bytes / 1024).toFixed(1)} KB
        </Caption1>
      </div>
      <div style={{
        display: 'flex', flexDirection: 'column', gap: 14, padding: 32,
        textAlign: 'center', alignItems: 'center',
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
      }}>
        <Icon style={{ width: 48, height: 48, color: tokens.colorPaletteRedForeground1 }} />
        <Text size={500} weight="semibold">{title}</Text>
        <Caption1 style={{ color: tokens.colorNeutralForeground2, maxWidth: 420 }}>
          {advice}
        </Caption1>
        {err && (
          // `whiteSpace: pre-wrap` keeps `\n` breaks but wraps the rest;
          // `wordBreak: break-word` + `overflowWrap: anywhere` break long
          // unbreakable strings (URLs, GUIDs, exception traces) so the
          // <pre> never expands past its container. `minWidth: 0` lets
          // it shrink inside the flex column.
          <pre style={{
            margin: 0, padding: 12,
            background: tokens.colorNeutralBackground1,
            border: `1px solid ${tokens.colorNeutralStroke2}`,
            borderRadius: tokens.borderRadiusMedium,
            fontFamily: tokens.fontFamilyMonospace, fontSize: 11,
            color: tokens.colorNeutralForeground1, textAlign: 'left',
            width: '100%', maxWidth: '100%', minWidth: 0,
            whiteSpace: 'pre-wrap',
            wordBreak: 'break-word',
            overflowWrap: 'anywhere',
            boxSizing: 'border-box',
          }}>{`code: ${err.code}\nmessage: ${err.message}`}</pre>
        )}
      </div>
    </div>
  );
}

// keep import warnings down
void Spinner; void ArrowDownload20Regular;
