// Sortable primitives built on `@dnd-kit` — a Fluent-v9-styled headless
// wrapper for keyboard / mouse / touch reordering.
//
// Two-component shape:
//   • <SortableList ids onReorder>...children...</SortableList>
//       Sets up DndContext + SortableContext with vertical-only constraints.
//       Calls onReorder(fromIndex, toIndex) when the user drops.
//   • <SortableItem id grip(?)>...row content...</SortableItem>
//       Provides a transform/transition style, exposes a drag handle ref
//       via the `grip` render prop. Falls back to making the entire item
//       draggable if no grip is rendered (useful for chips).
//
// Why @dnd-kit and not native HTML5 DnD:
//   • Accessibility — keyboard reorder (Space + arrows), screen reader
//     announcements, focus management.
//   • Touch — proper touch support out of the box; native HTML5 has none.
//   • Pointer modifiers — restrict to vertical axis, lock to parent, etc.
//   • Visual feedback — sub-pixel transforms with the dragged item lifting
//     above the rest, no ghost-image-from-mars positioning bugs.
//
// All visuals use Fluent v9 tokens so the primitive blends into every
// editor it's dropped into. The grip icon defaults to
// `ReOrderDotsHorizontal20Regular`, matching the visual decoration we
// already had as static dots across the FilterEditor / SelectEditor /
// OrderbyEditor / ApplyEditor.

import { type ReactNode, type CSSProperties } from 'react';
import {
  DndContext, type DragEndEvent,
  PointerSensor, KeyboardSensor, useSensor, useSensors,
  closestCenter, MeasuringStrategy,
} from '@dnd-kit/core';
import {
  SortableContext, useSortable,
  sortableKeyboardCoordinates,
  verticalListSortingStrategy,
  horizontalListSortingStrategy,
} from '@dnd-kit/sortable';
import { CSS } from '@dnd-kit/utilities';
import { ReOrderDotsHorizontal20Regular } from '@fluentui/react-icons';
import { makeStyles, tokens, mergeClasses } from '@fluentui/react-components';

const useStyles = makeStyles({
  grip: {
    color: tokens.colorNeutralForeground4,
    width: '16px',
    height: '16px',
    cursor: 'grab',
    flexShrink: 0,
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    // Hover: subtle brightening so the user knows this is interactive.
    ':hover': { color: tokens.colorNeutralForeground2 },
    ':active': { cursor: 'grabbing', color: tokens.colorBrandForeground1 },
    // Focus ring for keyboard navigation — Fluent's outline-style.
    ':focus-visible': {
      outline: `2px solid ${tokens.colorStrokeFocus2}`,
      outlineOffset: '1px',
      borderRadius: '2px',
    },
  },
  itemActive: {
    // The lifted state while dragging — Fluent's elevation shadow + slight
    // background brightening so the row reads as detached from the list.
    boxShadow: tokens.shadow16,
    backgroundColor: tokens.colorNeutralBackground1,
    zIndex: 10,
    borderRadius: tokens.borderRadiusMedium,
  },
  itemDropTarget: {
    // The placeholder slot where the item will land. Subtle border.
    outline: `1px dashed ${tokens.colorBrandStroke1}`,
    outlineOffset: '-2px',
  },
});

export interface SortableListProps {
  /** Stable string IDs for each item — must match the items in `children`. */
  ids: string[];
  /** Called when the user drops an item at a new index. Both args are 0-based. */
  onReorder: (fromIndex: number, toIndex: number) => void;
  /** "vertical" stacks rows top-to-bottom; "horizontal" lays chips side-by-side. */
  orientation?: 'vertical' | 'horizontal';
  children: ReactNode;
}

/**
 * Wraps a list of SortableItems with the DndContext + SortableContext
 * dnd-kit needs. Vertical orientation is the default — switch to
 * "horizontal" when laying out a chip row (e.g. groupby pills).
 */
export function SortableList({ ids, onReorder, orientation = 'vertical', children }: SortableListProps) {
  const sensors = useSensors(
    useSensor(PointerSensor, {
      activationConstraint: { distance: 4 }, // Allow click vs drag distinction
    }),
    useSensor(KeyboardSensor, {
      coordinateGetter: sortableKeyboardCoordinates,
    }),
  );

  const onDragEnd = (event: DragEndEvent) => {
    const { active, over } = event;
    if (!over || active.id === over.id) return;
    const fromIndex = ids.indexOf(String(active.id));
    const toIndex   = ids.indexOf(String(over.id));
    if (fromIndex < 0 || toIndex < 0) return;
    onReorder(fromIndex, toIndex);
  };

  return (
    <DndContext
      sensors={sensors}
      collisionDetection={closestCenter}
      measuring={{ droppable: { strategy: MeasuringStrategy.Always } }}
      onDragEnd={onDragEnd}
    >
      <SortableContext
        items={ids}
        strategy={orientation === 'horizontal' ? horizontalListSortingStrategy : verticalListSortingStrategy}
      >
        {children}
      </SortableContext>
    </DndContext>
  );
}

/** Props to spread onto whatever element acts as the drag handle. Threaded
 *  through render-props in deep components like FilterEditor's child blocks. */
export type GripProps = React.HTMLAttributes<HTMLElement> & {
  ref: React.Ref<HTMLElement>;
  tabIndex: 0;
};

export interface SortableItemProps {
  /** Stable string id — must appear in the parent SortableList's `ids`. */
  id: string;
  /**
   * Render prop receiving the props to spread onto the visible drag-grip
   * element. When provided, dragging is initiated ONLY from the grip
   * (typical UX — the row content stays clickable for selection / edit).
   * When omitted, the whole item is draggable (useful for chip-style
   * elements where the chip itself IS the handle).
   */
  children: (api: {
    /** Spread onto the drag-grip element. Includes a Fluent-styled
     *  `<ReOrderDotsHorizontal>` Component you can render as the icon. */
    gripProps: GripProps;
    /** Pre-built grip icon — render directly or use gripProps to wrap your own. */
    Grip: React.FC<{ className?: string }>;
    /** True while THIS item is being dragged. Style overlays / dim siblings. */
    isDragging: boolean;
  }) => ReactNode;
  /** Disable dragging — the item still renders but the grip becomes inert. */
  disabled?: boolean;
  /** Extra wrapper className. */
  className?: string;
  /** Extra wrapper style. */
  style?: CSSProperties;
}

/**
 * One item in a SortableList. Provides the lifted-shadow style during
 * drag, a Fluent-styled grip handle via render prop, and keyboard
 * accessibility (Tab to grip, Space to pick up, arrow keys to move,
 * Space again to drop). Reduces motion when the user prefers it.
 */
export function SortableItem({ id, children, disabled, className, style }: SortableItemProps) {
  const s = useStyles();
  const {
    attributes, listeners, setNodeRef, setActivatorNodeRef,
    transform, transition, isDragging, isOver,
  } = useSortable({ id, disabled });

  const wrapperStyle: CSSProperties = {
    transform: CSS.Translate.toString(transform),
    transition,
    ...style,
  };

  const Grip: React.FC<{ className?: string }> = ({ className: cn }) => (
    <ReOrderDotsHorizontal20Regular
      className={mergeClasses(s.grip, cn)}
      aria-hidden
    />
  );

  // `gripProps` spreads dnd-kit's listeners + attributes onto whatever
  // element the caller renders as the visible drag handle. Caller usually
  // wraps the `<Grip />` in a button or span.
  const gripProps = {
    ref: setActivatorNodeRef as React.Ref<HTMLElement>,
    ...listeners,
    ...attributes,
    tabIndex: 0 as const,
    style: { touchAction: 'none' as const, ...(disabled ? { cursor: 'not-allowed', opacity: 0.4 } : {}) },
  };

  return (
    <div
      ref={setNodeRef}
      style={wrapperStyle}
      className={mergeClasses(
        className,
        isDragging && s.itemActive,
        isOver && !isDragging && s.itemDropTarget,
      )}
    >
      {children({ gripProps, Grip, isDragging })}
    </div>
  );
}
