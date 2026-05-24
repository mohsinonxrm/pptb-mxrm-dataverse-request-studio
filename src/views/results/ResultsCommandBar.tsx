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
//   • Selection count + row count
//
// Search is intentionally NOT here — it's a column-list operation handled
// inside ResultsGrid. The command bar is for *actions on the result set*,
// not for filtering it.

import {
  Toolbar, ToolbarButton, ToolbarDivider, makeStyles, tokens, Tooltip,
  Menu, MenuTrigger, MenuPopover, MenuList, MenuItem, MenuButton, Spinner,
  Input, ToggleButton,
} from '@fluentui/react-components';
import {
  ArrowClockwise20Regular, Copy20Regular, Link20Regular,
  DocumentBulletList20Regular, ArrowDownload20Regular, Stack20Regular,
  Search20Regular, TextDensity20Regular, TextAlignDistributed20Regular,
} from '@fluentui/react-icons';
import { SegmentedToggle } from '../../primitives/SegmentedToggle';
import { useHostSession } from '../../host/HostContext';

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
  /** Row-search query (passed through to the grid). */
  searchQuery: string;
  onSearchChange: (q: string) => void;
  /** Row density. */
  density: 'cozy' | 'compact';
  onDensityChange: (d: 'cozy' | 'compact') => void;
}

export function ResultsCommandBar({
  hasMore, isLoading, isLoadingMore,
  requestUrl, body, onRefresh, onRetrieveAll,
  searchQuery, onSearchChange, density, onDensityChange,
}: ResultsCommandBarProps) {
  const s = useStyles();
  const host = useHostSession();
  // PPTB strips Prefer headers, so `odata.maxpagesize` for small-page-size
  // testing doesn't apply on Execute. When the response has no nextLink AND
  // we're inside PPTB, explain WHY pagination is unavailable rather than
  // just greying out the button.
  const retrieveAllTooltip = hasMore
    ? 'Keep fetching @odata.nextLink pages until everything is loaded'
    : host.embedded
      ? 'No more pages to fetch. Inside PPTB, Prefer: odata.maxpagesize is stripped, ' +
        'so the response uses Dataverse\'s default 5,000-row page size. To test paging at ' +
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
    const keys = Array.from(new Set(rows.flatMap(r => Object.keys(r))))
      .filter(k => !k.includes('@'));
    const header = keys.join('\t');
    const lines = rows.map(r =>
      keys.map(k => {
        const f = (r as Record<string, unknown>)[`${k}@OData.Community.Display.V1.FormattedValue`]
          ?? (r as Record<string, unknown>)[`_${k}_value@OData.Community.Display.V1.FormattedValue`];
        const v = f ?? r[k];
        if (v == null) return '';
        return String(v).replace(/[\t\n\r]+/g, ' ');
      }).join('\t'),
    );
    navigator.clipboard?.writeText([header, ...lines].join('\n'));
  };

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

        {onRetrieveAll && (
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

        <ToolbarDivider />

        <Input
          size="small"
          appearance="filled-lighter"
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
          >Cozy</ToggleButton>
          <ToggleButton
            checked={density === 'compact'}
            icon={<TextDensity20Regular />}
            onClick={() => onDensityChange('compact')}
            title="Compact"
          >Compact</ToggleButton>
        </SegmentedToggle>
      </Toolbar>

      <span className={s.spacer} />
    </div>
  );
}
