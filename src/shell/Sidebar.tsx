import type { FC, ReactNode } from 'react';
import { useStudioStyles } from '../primitives/styles';
import { ModeCard } from '../primitives/ModeCard';
import { SectionHeader } from '../primitives/SectionHeader';
import { ClauseTreeList, ClauseTreeItem } from '../primitives/ClauseTreeItem';
import { MethodPill } from '../primitives/MethodPill';
import type { RequestType } from '../registry/requestTypes';
import type { RecentRun } from '../state/readState';

export interface SidebarClauseItem {
  id: string; // selectable key
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  icon: FC<any>;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  iconFilled?: FC<any>;
  label: ReactNode;
  code?: boolean;
  badge?: ReactNode;
  badgeAppearance?: 'tint' | 'ghost' | 'filled' | 'outline';
  badgeColor?:
    | 'brand'
    | 'danger'
    | 'success'
    | 'warning'
    | 'informative'
    | 'subtle'
    | 'severe'
    | 'important';
  dirty?: boolean;
  /** Children rendered with depth=1 */
  children?: SidebarClauseItem[];
}

export interface SidebarSection {
  id: string;
  label: string;
  meta?: ReactNode;
  items: SidebarClauseItem[];
}

export interface SidebarProps {
  type: RequestType;
  urlPreview: string;
  sections: SidebarSection[];
  activeNode: string;
  onSelect: (id: string) => void;
  recents: RecentRun[];
  onRecentClick?: (r: RecentRun) => void;
}

export function Sidebar({
  type,
  urlPreview,
  sections,
  activeNode,
  onSelect,
  recents,
  onRecentClick,
}: SidebarProps) {
  const s = useStudioStyles();
  // A recent-run row is a button only when it can actually restore
  // something (#2): the mode wired a handler and the run has a snapshot.
  const isRestorable = (r: RecentRun) => !!onRecentClick && r.snapshot !== undefined;
  return (
    <aside className={s.sidebar}>
      <ModeCard type={type} urlPreview={urlPreview} />

      <div className={s.sidebarScroll}>
        {sections.map((sec) => (
          <span key={sec.id}>
            <SectionHeader meta={sec.meta}>{sec.label}</SectionHeader>
            <ClauseTreeList ariaLabel={sec.label}>
              {sec.items.map((item) => (
                <SidebarItemRenderer
                  key={item.id}
                  item={item}
                  activeNode={activeNode}
                  onSelect={onSelect}
                  group={type.group}
                  depth={0}
                />
              ))}
            </ClauseTreeList>
          </span>
        ))}

        {/* Hide the entire Recent runs section until there are runs. An
            empty-state ("No runs yet…") block adds ~50px of dead chrome
            on first paint of every mode; the section appears the first
            time the user clicks Execute. */}
        {recents.length > 0 && (
          <>
            <SectionHeader meta={`${recents.length} run${recents.length === 1 ? '' : 's'}`}>
              Recent runs
            </SectionHeader>
            <div style={{ padding: '4px 0' }}>
              {recents.map((r) => (
                <div
                  key={r.id}
                  className={isRestorable(r) ? s.recentRow : s.recentRowStatic}
                  role={isRestorable(r) ? 'button' : undefined}
                  tabIndex={isRestorable(r) ? 0 : undefined}
                  title={isRestorable(r) ? `Restore this run · ${r.url}` : r.url}
                  aria-label={
                    isRestorable(r)
                      ? `Restore run: ${r.method} ${r.url}, status ${r.status || 'error'}`
                      : undefined
                  }
                  onClick={() => isRestorable(r) && onRecentClick?.(r)}
                  onKeyDown={(e) => {
                    if (isRestorable(r) && (e.key === 'Enter' || e.key === ' ')) {
                      e.preventDefault();
                      onRecentClick?.(r);
                    }
                  }}
                >
                  <MethodPill method={r.method as never} size="sm" />
                  <span
                    style={{
                      flexGrow: 1,
                      overflow: 'hidden',
                      textOverflow: 'ellipsis',
                      whiteSpace: 'nowrap',
                    }}
                  >
                    {/* status 0 = the host failed the call without an HTTP status */}
                    {timeAgo(r.ts)} · {r.status || 'error'}
                    {runDetail(r)}
                  </span>
                </div>
              ))}
            </div>
          </>
        )}
      </div>
    </aside>
  );
}

function SidebarItemRenderer({
  item,
  activeNode,
  onSelect,
  group,
  depth,
}: {
  item: SidebarClauseItem;
  activeNode: string;
  onSelect: (id: string) => void;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  group: any;
  depth: 0 | 1 | 2 | 3;
}) {
  const expandable = !!item.children?.length;
  return (
    <span>
      <ClauseTreeItem
        icon={item.icon}
        iconFilled={item.iconFilled}
        label={item.label}
        code={item.code}
        selected={activeNode === item.id}
        dirty={item.dirty}
        badge={item.badge}
        badgeAppearance={item.badgeAppearance}
        badgeColor={item.badgeColor}
        expandable={expandable}
        expanded={expandable}
        onSelect={() => onSelect(item.id)}
        group={group}
        depth={depth}
      />
      {item.children?.map((child) => (
        <SidebarItemRenderer
          key={child.id}
          item={child}
          activeNode={activeNode}
          onSelect={onSelect}
          group={group}
          depth={Math.min(3, depth + 1) as 0 | 1 | 2 | 3}
        />
      ))}
    </span>
  );
}

function timeAgo(ts: number): string {
  const s = Math.floor((Date.now() - ts) / 1000);
  if (s < 60) return 'just now';
  const m = Math.floor(s / 60);
  if (m < 60) return `${m} min ago`;
  const h = Math.floor(m / 60);
  if (h < 24) return `${h} hr ago`;
  const d = Math.floor(h / 24);
  return `${d}d ago`;
}

/**
 * The tail of a Recent runs line: the mode's own summary, or a row count for
 * reads. A row count means nothing for a write ("204 · 1 rows"), so writes
 * without a summary show the status alone (#309).
 */
export function runDetail(r: Pick<RecentRun, 'method' | 'summary' | 'rowCount'>): string {
  if (r.summary) return ` · ${r.summary}`;
  if (r.method !== 'GET') return '';
  const n = r.rowCount;
  if (n === undefined) return ' · — rows';
  return ` · ${n.toLocaleString()} ${n === 1 ? 'row' : 'rows'}`;
}
