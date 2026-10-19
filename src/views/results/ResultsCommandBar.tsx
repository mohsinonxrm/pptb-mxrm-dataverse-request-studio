// ResultsCommandBar — top-of-grid toolbar for Retrieve Multiple results.
//
// Surface focuses on actions that make sense for an OData request studio
// (no Activate / Deactivate / Run Workflow / Server Export — those are
// model-driven-app affordances, out of scope here):
//
//   • Refresh      — re-execute the current query
//   • Retrieve all — keep fetching @odata.nextLink until exhausted
//   • Copy URL     — copy the resolved request URL
//   • Copy TSV     — copy current rows as tab-separated (paste-friendly)
//   • Copy JSON    — copy the OData envelope
//   • Export       — save CSV (as shown / raw) or the JSON envelope to a
//                    file via toolboxAPI.fileSystem.saveFile (#15)
//   • Selection count + row count
//
// Search is intentionally NOT here — it's a column-list operation handled
// inside ResultsGrid. The command bar is for *actions on the result set*,
// not for filtering it.

import {
  Toolbar,
  ToolbarButton,
  ToolbarDivider,
  makeStyles,
  tokens,
  Tooltip,
  Menu,
  MenuTrigger,
  MenuPopover,
  MenuList,
  MenuItem,
  MenuButton,
  Spinner,
  Input,
  ToggleButton,
  MenuDivider,
  MessageBar,
  MessageBarBody,
  MessageBarActions,
  Button,
} from '@fluentui/react-components';
import {
  ArrowClockwise20Regular,
  Copy20Regular,
  Link20Regular,
  DocumentBulletList20Regular,
  ArrowDownload20Regular,
  Stack20Regular,
  Search20Regular,
  TextDensity20Regular,
  TextAlignDistributed20Regular,
  ArrowExportUp20Regular,
  Dismiss20Regular,
} from '@fluentui/react-icons';
import { useState } from 'react';
import { SegmentedToggle } from '../../primitives/SegmentedToggle';
import { useHostSession } from '../../host/HostContext';
import { saveFileSafe } from '../../host/pptbBridge';
import { exportFileName, toCsv, type ExportValueMode, type GridExportView } from './exportResults';

const useStyles = makeStyles({
  bar: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    padding: '4px 8px',
    backgroundColor: tokens.colorNeutralBackground1,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    flexWrap: 'wrap',
  },
  meta: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    color: tokens.colorNeutralForeground3,
  },
  spacer: { flexGrow: 1 },
  notice: { flexBasis: '100%' },
});

export interface ResultsCommandBarProps {
  hasMore: boolean;
  isLoading: boolean;
  isLoadingMore: boolean;
  /** Current request URL (e.g. `/incidents?$select=…`) for "Copy URL". */
  requestUrl?: string;
  /** Raw OData envelope — used for "Copy JSON" / "Copy TSV". */
  body: unknown;
  /** Re-execute the current query. */
  onRefresh: () => void;
  /** Walk every page via @odata.nextLink until the result set is fully loaded. */
  onRetrieveAll?: () => void;
  /** Stop an in-flight page fetch; shown in place of Retrieve all while loading. */
  onStopLoadingMore?: () => void;
  /** Row-search query (passed through to the grid). */
  searchQuery: string;
  onSearchChange: (q: string) => void;
  /** Row density. */
  density: 'cozy' | 'compact';
  onDensityChange: (d: 'cozy' | 'compact') => void;
  /** Root table logical name — names the exported file. */
  table?: string;
  /** The columns + rows the grid lists; CSV export writes exactly these. */
  exportView?: GridExportView | null;
}

export function ResultsCommandBar({
  hasMore,
  isLoading,
  isLoadingMore,
  requestUrl,
  body,
  onRefresh,
  onRetrieveAll,
  onStopLoadingMore,
  searchQuery,
  onSearchChange,
  density,
  onDensityChange,
  table,
  exportView,
}: ResultsCommandBarProps) {
  const s = useStyles();
  const host = useHostSession();
  const [exporting, setExporting] = useState(false);
  const [notice, setNotice] = useState<{ intent: 'success' | 'error'; text: string } | null>(null);
  // PPTB strips Prefer headers, so `odata.maxpagesize` for small-page-size
  // testing doesn't apply on Execute. When the response has no nextLink AND
  // we're inside PPTB, explain WHY pagination is unavailable rather than
  // just greying out the button.
  const retrieveAllTooltip = hasMore
    ? 'Keep fetching @odata.nextLink pages until everything is loaded'
    : host.embedded
      ? 'No more pages to fetch. Inside PPTB, Prefer: odata.maxpagesize is stripped, ' +
        "so the response uses Dataverse's default 5,000-row page size. To test paging at " +
        'smaller sizes, export the request via the Code tab.'
      : 'No more pages to fetch — the response did not include @odata.nextLink. ' +
        'Set Prefer: odata.maxpagesize to enable server-side pagination.';

  const copyUrl = () => {
    if (requestUrl) navigator.clipboard?.writeText(requestUrl);
  };

  const copyJson = () => {
    navigator.clipboard?.writeText(JSON.stringify(body, null, 2));
  };

  const copyTsv = () => {
    const rows = (body as { value?: Record<string, unknown>[] } | null)?.value ?? [];
    if (rows.length === 0) return;
    const keys = Array.from(new Set(rows.flatMap((r) => Object.keys(r)))).filter(
      (k) => !k.includes('@'),
    );
    const header = keys.join('\t');
    const lines = rows.map((r) =>
      keys
        .map((k) => {
          const f =
            (r as Record<string, unknown>)[`${k}@OData.Community.Display.V1.FormattedValue`] ??
            (r as Record<string, unknown>)[`_${k}_value@OData.Community.Display.V1.FormattedValue`];
          const v = f ?? r[k];
          if (v == null) return '';
          return String(v).replace(/[\t\n\r]+/g, ' ');
        })
        .join('\t'),
    );
    navigator.clipboard?.writeText([header, ...lines].join('\n'));
  };

  const save = async (fileName: string, content: string, filter: { name: string; ext: string }) => {
    setExporting(true);
    setNotice(null);
    try {
      const path = await saveFileSafe(fileName, content, [
        { name: filter.name, extensions: [filter.ext] },
      ]);
      // PPTB returns the chosen path (null/undefined when the dialog was
      // cancelled); the standalone browser download returns null too.
      if (path) setNotice({ intent: 'success', text: `Saved to ${path}` });
    } catch (e) {
      setNotice({
        intent: 'error',
        text: `Export failed: ${e instanceof Error ? e.message : String(e)}`,
      });
    } finally {
      setExporting(false);
    }
  };

  const exportCsv = (mode: ExportValueMode) => {
    if (!exportView || exportView.rows.length === 0) return;
    void save(exportFileName(table, 'csv'), toCsv(exportView.rows, exportView.columns, mode), {
      name: 'CSV',
      ext: 'csv',
    });
  };

  const exportJson = () => {
    void save(exportFileName(table, 'json'), JSON.stringify(body, null, 2), {
      name: 'JSON',
      ext: 'json',
    });
  };

  const csvDisabled = exporting || !exportView || exportView.rows.length === 0;

  return (
    <div className={s.bar}>
      <Toolbar size="small" aria-label="Results actions">
        <Tooltip content="Re-run the request" relationship="label">
          <ToolbarButton
            appearance="subtle"
            icon={isLoading ? <Spinner size="tiny" /> : <ArrowClockwise20Regular />}
            disabled={isLoading}
            onClick={onRefresh}
          >
            Refresh
          </ToolbarButton>
        </Tooltip>

        {onRetrieveAll && isLoadingMore && onStopLoadingMore ? (
          <Tooltip
            content="Stop fetching pages. Rows loaded so far are kept."
            relationship="description"
          >
            <ToolbarButton
              appearance="subtle"
              icon={<Spinner size="tiny" />}
              onClick={onStopLoadingMore}
            >
              Stop
            </ToolbarButton>
          </Tooltip>
        ) : (
          onRetrieveAll && (
            <Tooltip content={retrieveAllTooltip} relationship="description">
              <ToolbarButton
                appearance="subtle"
                icon={isLoadingMore ? <Spinner size="tiny" /> : <Stack20Regular />}
                disabled={!hasMore || isLoadingMore}
                onClick={onRetrieveAll}
              >
                Retrieve all
              </ToolbarButton>
            </Tooltip>
          )
        )}

        <ToolbarDivider />

        <Tooltip content="Copy the resolved request URL" relationship="label">
          <ToolbarButton
            appearance="subtle"
            icon={<Link20Regular />}
            disabled={!requestUrl}
            onClick={copyUrl}
          >
            Copy URL
          </ToolbarButton>
        </Tooltip>

        <Menu>
          <MenuTrigger disableButtonEnhancement>
            <MenuButton appearance="subtle" size="small" icon={<ArrowDownload20Regular />}>
              Copy data
            </MenuButton>
          </MenuTrigger>
          <MenuPopover>
            <MenuList>
              <MenuItem icon={<DocumentBulletList20Regular />} onClick={copyTsv}>
                Copy as TSV (Excel-friendly)
              </MenuItem>
              <MenuItem icon={<Copy20Regular />} onClick={copyJson}>
                Copy JSON envelope
              </MenuItem>
            </MenuList>
          </MenuPopover>
        </Menu>

        <Menu>
          <MenuTrigger disableButtonEnhancement>
            <MenuButton
              appearance="subtle"
              size="small"
              icon={exporting ? <Spinner size="tiny" /> : <ArrowExportUp20Regular />}
              disabled={exporting}
            >
              Export
            </MenuButton>
          </MenuTrigger>
          <MenuPopover>
            <MenuList>
              <MenuItem disabled={csvDisabled} onClick={() => exportCsv('shown')}>
                CSV — values as shown
              </MenuItem>
              <MenuItem disabled={csvDisabled} onClick={() => exportCsv('raw')}>
                CSV — raw values (GUIDs, numbers, ISO dates)
              </MenuItem>
              <MenuDivider />
              <MenuItem disabled={exporting} onClick={exportJson}>
                JSON — OData envelope
              </MenuItem>
            </MenuList>
          </MenuPopover>
        </Menu>

        <ToolbarDivider />

        <Input
          size="small"
          appearance="filled-lighter"
          aria-label="Search rows"
          placeholder="Search rows…"
          contentBefore={<Search20Regular />}
          value={searchQuery}
          onChange={(_, d) => onSearchChange(d.value)}
          style={{ minWidth: 220 }}
        />

        <SegmentedToggle ariaLabel="Row density">
          <ToggleButton
            checked={density === 'cozy'}
            icon={<TextAlignDistributed20Regular />}
            onClick={() => onDensityChange('cozy')}
            title="Cozy"
          >
            Cozy
          </ToggleButton>
          <ToggleButton
            checked={density === 'compact'}
            icon={<TextDensity20Regular />}
            onClick={() => onDensityChange('compact')}
            title="Compact"
          >
            Compact
          </ToggleButton>
        </SegmentedToggle>
      </Toolbar>

      <span className={s.spacer} />

      {notice && (
        <MessageBar intent={notice.intent} className={s.notice}>
          <MessageBarBody>{notice.text}</MessageBarBody>
          <MessageBarActions
            containerAction={
              <Button
                appearance="transparent"
                aria-label="Dismiss"
                icon={<Dismiss20Regular />}
                onClick={() => setNotice(null)}
              />
            }
          />
        </MessageBar>
      )}
    </div>
  );
}
