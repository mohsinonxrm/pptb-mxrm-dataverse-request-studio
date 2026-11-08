import type { FC, ReactNode } from 'react';
import { Badge, mergeClasses } from '@fluentui/react-components';
import { Checkmark16Filled } from '@fluentui/react-icons';
import { useStudioStyles } from './styles';
import { groupColorVar } from '../theme/theme';
import type { RequestGroup } from '../registry/requestTypes';

// One clause in the sidebar (#243): a button in a plain list, inside the
// sidebar's <nav>. Tab moves between clauses, Enter or Space opens one, and
// the open one has aria-current. Nested clauses ($expand clauses, required
// parameters, Bulk Write rows) sit in a nested list, so screen readers
// announce their level.
// The sidebar used to claim role="tree" without the tree keyboard model,
// and its chevrons looked collapsible, but nothing ever collapsed.
// Owns: icon / label / aside (badge) / dirty marker / click handler.
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
  badgeColor?:
    | 'brand'
    | 'danger'
    | 'success'
    | 'warning'
    | 'informative'
    | 'subtle'
    | 'severe'
    | 'important';
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
      aria-current={selected ? 'true' : undefined}
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
      {/* First grid column, where the chevron used to be; keeps the icons aligned. */}
      <span aria-hidden />
      <ResolvedIcon
        className={s.ctiIcon}
        style={{ color: selected ? groupColorVar(group) : undefined, opacity: selected ? 1 : 0.85 }}
      />
      <span className={mergeClasses(s.ctiLabel, code && s.ctiCodeLabel)}>{label}</span>
      <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6, justifySelf: 'end' }}>
        {dirty && (
          <>
            <span className={s.ctiDirty} aria-hidden>
              •
            </span>
            {/* A plain span can't carry aria-label, so say it in text. */}
            <span className={s.srOnly}>(modified)</span>
          </>
        )}
        {badge !== undefined &&
          badge !== null &&
          badge !== false &&
          /* Replace the literal '✓' string badge (reads as data, not state)
             with an iconic Checkmark in a success-tinted badge — handled
             here so callsites don't each have to import the icon. */
          (badge === '✓' ? (
            <Badge
              size="small"
              appearance="tint"
              color="success"
              icon={<Checkmark16Filled />}
              role="img"
              aria-label="Set"
            />
          ) : (
            // `small` with an 11px label (#200): Fluent's `extra-small` is a
            // 6px dot and squeezed text badges to ~4px. Long labels
            // ("PATCH 1 field") truncate instead of pushing the row wider.
            <Badge
              size="small"
              appearance={badgeAppearance}
              color={badgeColor}
              className={s.ctiBadge}
              title={typeof badge === 'string' ? badge : undefined}
            >
              {badge}
            </Badge>
          ))}
      </span>
    </button>
  );
}

/** A list of clauses: a sidebar section, or the clauses nested under one. */
export function ClauseTreeList({
  children,
  labelledBy,
}: {
  children: ReactNode;
  /** Id of the section heading that names the list. */
  labelledBy?: string;
}) {
  const s = useStudioStyles();
  return (
    <ul aria-labelledby={labelledBy} className={s.ctiList}>
      {children}
    </ul>
  );
}
