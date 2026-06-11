import type { FC, ReactNode } from 'react';
import { Badge, mergeClasses } from '@fluentui/react-components';
import {
  ChevronRight20Regular,
  ChevronDown20Regular,
  Checkmark16Filled,
} from '@fluentui/react-icons';
import { useStudioStyles } from './styles';
import { groupColorVar } from '../theme/theme';
import type { RequestGroup } from '../registry/requestTypes';

// Replaces Fluent's TreeItemLayout entirely. Owns:
// chev / icon / label / aside (badge) / dirty marker / click handler.
export interface ClauseTreeItemProps {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  icon: FC<any>;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  iconFilled?: FC<any>;
  label: ReactNode;
  /** Render label in monospace */
  code?: boolean;
  /** Active editor selects this row */
  selected?: boolean;
  /** User has changed this clause since last execute */
  dirty?: boolean;
  /** Right-aligned badge (count or sub) */
  badge?: ReactNode;
  badgeAppearance?: 'tint' | 'ghost' | 'filled' | 'outline';
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  badgeColor?:
    | 'brand'
    | 'danger'
    | 'success'
    | 'warning'
    | 'informative'
    | 'subtle'
    | 'severe'
    | 'important';
  /** Whether this row expands children */
  expandable?: boolean;
  expanded?: boolean;
  onToggle?: () => void;
  /** Click swaps the right pane */
  onSelect?: () => void;
  /** Indent depth (0 default, 1+ for nested children) */
  depth?: 0 | 1 | 2 | 3;
  /** Tint color for selected state — defaults to read group */
  group?: RequestGroup;
}

export function ClauseTreeItem(props: ClauseTreeItemProps) {
  const s = useStudioStyles();
  const {
    icon: Icon,
    iconFilled: IconFilled,
    label,
    code,
    selected,
    dirty,
    badge,
    badgeAppearance = 'tint',
    badgeColor = 'brand',
    expandable,
    expanded,
    onToggle,
    onSelect,
    depth = 0,
    group = 'read',
  } = props;
  const ResolvedIcon = selected && IconFilled ? IconFilled : Icon;
  const indentClass =
    depth === 1
      ? s.ctiIndent1
      : depth === 2
        ? s.ctiIndent2
        : depth === 3
          ? s.ctiIndent3
          : undefined;

  return (
    <button
      type="button"
      role="treeitem"
      aria-selected={selected || undefined}
      aria-expanded={expandable ? !!expanded : undefined}
      className={mergeClasses(s.cti, indentClass, selected && s.ctiSelected)}
      onClick={onSelect}
    >
      {selected && (
        <span
          className={s.ctiSelectedRule}
          style={{ background: groupColorVar(group) }}
          aria-hidden
        />
      )}
      <span
        onClick={(e) => {
          if (expandable && onToggle) {
            e.stopPropagation();
            onToggle();
          }
        }}
        style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }}
      >
        {expandable ? (
          expanded ? (
            <ChevronDown20Regular className={s.ctiChev} />
          ) : (
            <ChevronRight20Regular className={s.ctiChev} />
          )
        ) : (
          <span style={{ width: 14 }} aria-hidden />
        )}
      </span>
      <ResolvedIcon
        className={s.ctiIcon}
        style={{ color: selected ? groupColorVar(group) : undefined, opacity: selected ? 1 : 0.85 }}
      />
      <span className={mergeClasses(s.ctiLabel, code && s.ctiCodeLabel)}>{label}</span>
      <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6, justifySelf: 'end' }}>
        {dirty && (
          <span className={s.ctiDirty} aria-label="Modified">
            •
          </span>
        )}
        {badge !== undefined &&
          badge !== null &&
          badge !== false &&
          /* Replace the literal '✓' string badge (reads as data, not state)
             with an iconic Checkmark in a success-tinted badge — handled
             here so callsites don't each have to import the icon. */
          (badge === '✓' ? (
            <Badge
              size="extra-small"
              appearance="tint"
              color="success"
              icon={<Checkmark16Filled />}
              aria-label="Set"
            />
          ) : (
            <Badge size="extra-small" appearance={badgeAppearance} color={badgeColor}>
              {badge}
            </Badge>
          ))}
      </span>
    </button>
  );
}

export function ClauseTreeList({
  children,
  ariaLabel,
}: {
  children: ReactNode;
  ariaLabel?: string;
}) {
  const s = useStudioStyles();
  return (
    <ul role="tree" aria-label={ariaLabel} className={s.ctiList}>
      {children}
    </ul>
  );
}
