import { useId, useState, type FC, type ReactNode } from 'react';
import { Button } from '@fluentui/react-components';
import { History16Regular } from '@fluentui/react-icons';
import { useStudioStyles } from '../primitives/styles';
import { ModeCard } from '../primitives/ModeCard';
import { SectionHeader } from '../primitives/SectionHeader';
import { ClauseTreeList, ClauseTreeItem } from '../primitives/ClauseTreeItem';
import { MethodPill } from '../primitives/MethodPill';
import type { RequestType } from '../registry/requestTypes';
import type { RecentRun } from '../state/readState';
import { useModeHistory } from '../state/runHistory';
import { RunHistoryDialog } from './RunHistoryDialog';
import { runDetail, timeAgo } from './runLabels';
import { notify } from './dialogs';

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
  /** Clauses nested under this one: a nested list, one level deeper */
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
  const uid = useId();
  // A recent-run row is a button only when it can actually restore
  // something (#2): the mode wired a handler and the run has a snapshot.
  const isRestorable = (r: RecentRun) => !!onRecentClick && r.snapshot !== undefined;
  // A kept run had its caller id removed (#226): say so, since the restored
  // request now runs as the signed-in user.
  const restore = (r: RecentRun) => {
    onRecentClick?.(r);
    if (r.impersonationCleared) {
      notify(
        'This run impersonated another user. History doesn’t keep the caller id, so the restored request runs as you: set MSCRMCallerID or CallerObjectId again in Headers to impersonate.',
        { title: 'Impersonation not restored' },
      );
    }
  };
  // Every run kept for this mode and environment (#226); the sidebar lists the latest.
  const history = useModeHistory(type.id, recents);
  const [historyOpen, setHistoryOpen] = useState(false);
  return (
    <aside className={s.sidebar}>
      <ModeCard type={type} urlPreview={urlPreview} />

      <div className={s.sidebarScroll}>
        {/* The clauses (#243): a list of buttons per section, each list
            named by its heading. Nested clauses are nested lists. */}
        <nav aria-label="Request structure">
          {sections.map((sec) => (
            <div key={sec.id}>
              <SectionHeader id={`${uid}-${sec.id}`} meta={sec.meta}>
                {sec.label}
              </SectionHeader>
              <ClauseTreeList labelledBy={`${uid}-${sec.id}`}>
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
            </div>
          ))}
        </nav>

        {/* Hide the entire Recent runs section until there are runs. An
            empty-state ("No runs yet…") block adds ~50px of dead chrome
            on first paint of every mode; the section appears the first
            time the user clicks Execute. */}
        {recents.length > 0 && (
          <>
            <SectionHeader
              meta={
                history.length > recents.length
                  ? `${recents.length} of ${history.length}`
                  : `${recents.length} run${recents.length === 1 ? '' : 's'}`
              }
            >
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
                  onClick={() => isRestorable(r) && restore(r)}
                  onKeyDown={(e) => {
                    if (isRestorable(r) && (e.key === 'Enter' || e.key === ' ')) {
                      e.preventDefault();
                      restore(r);
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
              <Button
                appearance="subtle"
                size="small"
                icon={<History16Regular />}
                onClick={() => setHistoryOpen(true)}
                style={{ marginTop: 2 }}
              >
                {history.length > recents.length
                  ? `Show all ${history.length} runs`
                  : 'Search runs'}
              </Button>
            </div>
          </>
        )}
      </div>
      <RunHistoryDialog
        open={historyOpen}
        onOpenChange={setHistoryOpen}
        modeName={type.name}
        runs={history}
        onRestore={onRecentClick ? restore : undefined}
      />
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
  return (
    <li>
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
        onSelect={() => onSelect(item.id)}
        group={group}
        depth={depth}
      />
      {!!item.children?.length && (
        <ClauseTreeList>
          {item.children.map((child) => (
            <SidebarItemRenderer
              key={child.id}
              item={child}
              activeNode={activeNode}
              onSelect={onSelect}
              group={group}
              depth={Math.min(3, depth + 1) as 0 | 1 | 2 | 3}
            />
          ))}
        </ClauseTreeList>
      )}
    </li>
  );
}
