// RecordDetailCard — recursive renderer for a single Dataverse record.
//
// Design (decided with the user):
//   • The user came to verify a response, not to navigate a record.
//     Inline accordions per nav property keep one scrollable surface.
//   • Nested expands recurse — depth is whatever the server returned.
//     We don't lazy-load and we don't add "open the related record"
//     navigation. Render what came back, period.
//   • N:1 / 1:1 expansions render as a sub-card (recursive RecordDetailCard).
//   • 1:N / N:N expansions render as a CollectionSubgrid (Fluent v9 Table)
//     with a per-row toggle that swaps the row in for a full sub-card —
//     so a 1:N → N:1 chain is still expandable end-to-end.
//   • OData annotations (@odata.etag, @odata.context, etc.) get tucked
//     into a collapsed panel at the bottom — the casual reader skips
//     them, the curious developer expands.
//
// Fluent v9 components used:
//   Accordion / AccordionItem / AccordionHeader / AccordionPanel  (open/close)
//   Card / CardHeader                                              (visual shell)
//   Badge                                                          (N:1 / 1:N labels)
//   Persona                                                        (headline strip)
//   Body1 / Caption1                                               (typography)

import {
  Accordion,
  AccordionItem,
  AccordionHeader,
  AccordionPanel,
  Badge,
  Body1,
  Caption1,
  Persona,
  tokens,
  CounterBadge,
} from '@fluentui/react-components';
import {
  CheckmarkCircle20Filled,
  LinkSquare16Regular,
  PuzzlePiece16Regular,
  BranchFork16Regular,
  BranchFork20Regular,
} from '@fluentui/react-icons';
import { findTable } from '../../mock/metadata';
import {
  partitionRecord,
  prettifyKey,
  pickHeadlineFields,
  type RecordPartition,
  type PartitionedScalar,
} from './detailFieldPartitioner';
import { CollectionSubgrid } from './CollectionSubgrid';
import { ODataAnnotationsPanel } from './ODataAnnotationsPanel';

export interface RecordDetailCardProps {
  /** The record to render. */
  record: Record<string, unknown>;
  /** Logical name of the table this record belongs to. Drives nav-property
   *  metadata lookup. Empty / unknown is fine — partitioner falls back to
   *  inference and we render slightly less context. */
  entityLogical: string;
  /** Recursion depth. Drives default-open vs default-closed accordions
   *  and the headline persona density. */
  level?: number;
  /** Render the headline persona strip at the top. Off for sub-cards
   *  rendered inside an accordion panel (the accordion header IS the
   *  headline there). */
  showHeadline?: boolean;
}

const DEFAULT_OPEN_OBJECTS_UP_TO_LEVEL = 1; // N:1 cards open through level 1
const DEFAULT_OPEN_COLLECTIONS_UP_TO_LEVEL = 0; // 1:N grids open only at root
const COLLECTION_AUTO_COLLAPSE_THRESHOLD = 5; // collapse 1:N if more than N rows

export function RecordDetailCard({
  record,
  entityLogical,
  level = 0,
  showHeadline = true,
}: RecordDetailCardProps) {
  const tbl = findTable(entityLogical);
  const part: RecordPartition = partitionRecord(record, tbl);
  const { headline, subline } = pickHeadlineFields(record, tbl);

  return (
    <div
      style={{
        border: level === 0 ? `1px solid ${tokens.colorNeutralStroke2}` : 'none',
        borderRadius: tokens.borderRadiusLarge,
        padding: level === 0 ? 20 : 0,
        background: level === 0 ? tokens.colorNeutralBackground1 : 'transparent',
      }}
    >
      {showHeadline && level === 0 && (
        <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 16 }}>
          <Persona
            name={headline}
            primaryText={headline}
            secondaryText={tbl?.displayName ?? entityLogical ?? subline}
            tertiaryText={subline}
            size="large"
            avatar={{ color: 'colorful' }}
          />
          <span style={{ flexGrow: 1 }} />
          <CheckmarkCircle20Filled style={{ color: tokens.colorPaletteGreenForeground1 }} />
        </div>
      )}

      <ScalarGrid scalars={part.scalars} />

      {/* N:1 / 1:1 expanded nav objects → recursive sub-card */}
      {part.navObjects.length > 0 && (
        <Accordion
          multiple
          collapsible
          defaultOpenItems={
            level <= DEFAULT_OPEN_OBJECTS_UP_TO_LEVEL
              ? part.navObjects.map((n) => `obj-${n.key}`)
              : []
          }
          style={{ marginTop: 16 }}
        >
          {part.navObjects.map((n) => (
            <AccordionItem key={`obj-${n.key}`} value={`obj-${n.key}`}>
              <AccordionHeader expandIconPosition="end" icon={<LinkSquare16Regular />}>
                <NavHeader
                  navKey={n.key}
                  targetDisplay={n.targetEntityDisplay}
                  targetLogical={n.targetEntityLogical}
                  cardinality={n.cardinality}
                />
              </AccordionHeader>
              <AccordionPanel>
                <div
                  style={{
                    marginLeft: 8,
                    paddingLeft: 12,
                    borderLeft: `2px solid ${tokens.colorNeutralStroke3}`,
                  }}
                >
                  <RecordDetailCard
                    record={n.value}
                    entityLogical={n.targetEntityLogical ?? ''}
                    level={level + 1}
                    showHeadline={false}
                  />
                </div>
              </AccordionPanel>
            </AccordionItem>
          ))}
        </Accordion>
      )}

      {/* 1:N / N:N expanded collections → sub-grid */}
      {part.navCollections.length > 0 && (
        <Accordion
          multiple
          collapsible
          defaultOpenItems={
            level <= DEFAULT_OPEN_COLLECTIONS_UP_TO_LEVEL
              ? part.navCollections
                  .filter((n) => n.value.length <= COLLECTION_AUTO_COLLAPSE_THRESHOLD)
                  .map((n) => `col-${n.key}`)
              : []
          }
          style={{ marginTop: 12 }}
        >
          {part.navCollections.map((n) => (
            <AccordionItem key={`col-${n.key}`} value={`col-${n.key}`}>
              <AccordionHeader expandIconPosition="end" icon={<BranchFork16Regular />}>
                <NavHeader
                  navKey={n.key}
                  targetDisplay={n.targetEntityDisplay}
                  targetLogical={n.targetEntityLogical}
                  cardinality={n.cardinality}
                  count={n.value.length}
                />
              </AccordionHeader>
              <AccordionPanel>
                <div
                  style={{
                    marginLeft: 8,
                    paddingLeft: 12,
                    borderLeft: `2px solid ${tokens.colorNeutralStroke3}`,
                  }}
                >
                  {n.value.length === 0 ? (
                    <Caption1
                      style={{ color: tokens.colorNeutralForeground3, fontStyle: 'italic' }}
                    >
                      (no related rows)
                    </Caption1>
                  ) : (
                    <CollectionSubgrid
                      rows={n.value}
                      entityLogical={n.targetEntityLogical ?? ''}
                      level={level + 1}
                    />
                  )}
                </div>
              </AccordionPanel>
            </AccordionItem>
          ))}
        </Accordion>
      )}

      {/* Null-expanded nav properties — make them visible so the user
          knows "this WAS queried, came back unset" rather than "absent" */}
      {part.navNulls.length > 0 && (
        <div style={{ marginTop: 12, display: 'flex', flexDirection: 'column', gap: 4 }}>
          {part.navNulls.map((n) => (
            <div
              key={`null-${n.key}`}
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: 8,
                padding: '6px 10px',
                border: `1px dashed ${tokens.colorNeutralStroke2}`,
                borderRadius: tokens.borderRadiusMedium,
                color: tokens.colorNeutralForeground3,
                fontSize: tokens.fontSizeBase200,
              }}
            >
              <BranchFork20Regular style={{ opacity: 0.6 }} />
              <span style={{ fontFamily: tokens.fontFamilyMonospace }}>{n.key}</span>
              <Caption1 style={{ color: tokens.colorNeutralForeground4 }}>
                → {n.targetEntityDisplay ?? n.targetEntityLogical ?? '?'} ·
              </Caption1>
              <Caption1 style={{ color: tokens.colorNeutralForeground4, fontStyle: 'italic' }}>
                (not set)
              </Caption1>
            </div>
          ))}
        </div>
      )}

      {/* OData metadata annotations — collapsed by default */}
      {part.annotations.length > 0 && level === 0 && (
        <ODataAnnotationsPanel annotations={part.annotations} />
      )}
    </div>
  );
}

// ── Scalar grid ──────────────────────────────────────────────────────

function ScalarGrid({ scalars }: { scalars: PartitionedScalar[] }) {
  if (scalars.length === 0) return null;
  return (
    <div
      style={{
        display: 'grid',
        // Key column wraps if the field name is long; value column flexes.
        gridTemplateColumns: 'minmax(120px, 200px) 1fr',
        rowGap: 8,
        columnGap: 16,
        alignItems: 'baseline',
      }}
    >
      {scalars.map((s) => (
        <ScalarRow key={s.key} s={s} />
      ))}
    </div>
  );
}

function ScalarRow({ s }: { s: PartitionedScalar }) {
  const isLookupWireForm = /^_(.+)_value$/.test(s.key);
  const displayKey = prettifyKey(s.key);
  return (
    <>
      <span
        style={{
          color: tokens.colorNeutralForeground3,
          fontSize: tokens.fontSizeBase200,
          fontWeight: tokens.fontWeightSemibold,
          textTransform: 'uppercase',
          letterSpacing: '0.02em',
          fontFamily: tokens.fontFamilyBase,
          wordBreak: 'break-word',
        }}
        title={s.key}
      >
        {displayKey}
        {isLookupWireForm && (
          <Badge appearance="ghost" size="extra-small" style={{ marginLeft: 6 }}>
            lookup
          </Badge>
        )}
      </span>
      <span style={{ wordBreak: 'break-word', fontSize: tokens.fontSizeBase300 }}>
        <ScalarValue s={s} />
      </span>
    </>
  );
}

function ScalarValue({ s }: { s: PartitionedScalar }) {
  if (s.value == null) {
    return <span style={{ color: tokens.colorNeutralForeground4 }}>—</span>;
  }
  // Formatted value takes priority. Show the raw underneath in a smaller
  // mono caption so the user still knows the underlying GUID / option-set
  // value / state code without having to click the JSON tab.
  if (s.formattedValue != null) {
    return (
      <span>
        <Body1>{s.formattedValue}</Body1>
        <span
          style={{
            marginLeft: 6,
            fontFamily: tokens.fontFamilyMonospace,
            fontSize: 11,
            color: tokens.colorNeutralForeground3,
          }}
        >
          ({String(s.value)})
        </span>
        {s.lookupTargetLogicalName && (
          <Badge
            appearance="tint"
            size="extra-small"
            icon={<PuzzlePiece16Regular />}
            style={{ marginLeft: 6, verticalAlign: 'middle' }}
          >
            {s.lookupTargetLogicalName}
          </Badge>
        )}
      </span>
    );
  }
  // No formatted value — print the raw. Mono-font for GUIDs so they're
  // recognizable at a glance.
  const raw = s.value;
  const isStringy = typeof raw === 'string';
  const looksLikeGuid =
    isStringy &&
    /^[{(]?[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}[)}]?$/.test(
      raw as string,
    );
  return (
    <span
      style={{
        fontFamily: looksLikeGuid ? tokens.fontFamilyMonospace : tokens.fontFamilyBase,
        fontSize: looksLikeGuid ? 12 : tokens.fontSizeBase300,
      }}
    >
      {String(raw)}
      {s.lookupTargetLogicalName && (
        <Badge
          appearance="tint"
          size="extra-small"
          icon={<PuzzlePiece16Regular />}
          style={{ marginLeft: 6, verticalAlign: 'middle' }}
        >
          {s.lookupTargetLogicalName}
        </Badge>
      )}
    </span>
  );
}

// ── Accordion header content ─────────────────────────────────────────

function NavHeader({
  navKey,
  targetDisplay,
  targetLogical,
  cardinality,
  count,
}: {
  navKey: string;
  targetDisplay?: string;
  targetLogical?: string;
  cardinality?: 'OneToMany' | 'ManyToOne' | 'ManyToMany';
  count?: number;
}) {
  const cardLabel =
    cardinality === 'OneToMany'
      ? '1:N'
      : cardinality === 'ManyToOne'
        ? 'N:1'
        : cardinality === 'ManyToMany'
          ? 'N:N'
          : null;
  const cardColor: 'brand' | 'informative' | 'subtle' =
    cardinality === 'ManyToOne'
      ? 'informative'
      : cardinality === 'OneToMany' || cardinality === 'ManyToMany'
        ? 'brand'
        : 'subtle';
  return (
    <span style={{ display: 'inline-flex', alignItems: 'center', gap: 8, minWidth: 0 }}>
      <span style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: tokens.fontSizeBase300 }}>
        {navKey}
      </span>
      {(targetDisplay || targetLogical) && (
        <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
          → {targetDisplay || targetLogical}
        </Caption1>
      )}
      {cardLabel && (
        <Badge appearance="tint" color={cardColor} size="small">
          {cardLabel}
        </Badge>
      )}
      {typeof count === 'number' && (
        <CounterBadge appearance="filled" count={count} overflowCount={9999} size="small" />
      )}
    </span>
  );
}
