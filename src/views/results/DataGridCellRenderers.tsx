// Type-aware cell renderers — one per AttributeTypeCode, each preferring
// the OData formatted-value annotation when present and falling back to
// typed formatting otherwise.
//
// Choose the right Fluent v9 surface per type so the grid reads like
// Model-Driven Apps — checkboxes for Boolean, formatted currency for Money,
// chips for OptionSets, etc. Each renderer dispatches on `ColumnMeta`'s
// `attributeType` discriminator and falls back to a plain text rendering
// when the value is missing or unrecognized.

import { Switch, Badge, Label, makeStyles, tokens } from '@fluentui/react-components';
import type {
  ColumnMeta,
  BooleanColumnMeta,
  MoneyColumnMeta,
  DecimalColumnMeta,
  DoubleColumnMeta,
  DateTimeColumnMeta,
  PicklistColumnMeta,
  StatusColumnMeta,
  StateColumnMeta,
  MultiSelectPicklistColumnMeta,
} from '../../mock/metadata';

const useStyles = makeStyles({
  booleanCell: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
  },
  lookupCell: {
    display: 'flex',
    flexDirection: 'column',
    gap: '2px',
    minWidth: 0,
  },
  lookupName: {
    display: 'block',
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  lookupType: {
    display: 'block',
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  numberCell: {
    display: 'block',
    textAlign: 'right',
    fontVariantNumeric: 'tabular-nums',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  dateCell: {
    display: 'block',
    fontVariantNumeric: 'tabular-nums',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  picklistCell: {
    display: 'flex',
    alignItems: 'center',
  },
  truncateText: {
    display: 'block',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  emptyDash: {
    color: tokens.colorNeutralForeground4,
  },
});

const EmptyDash = () => {
  const s = useStyles();
  return <span className={s.emptyDash}>—</span>;
};

/** Boolean → disabled Switch + label. Formatted value preferred for the label. */
function BooleanCellRenderer({
  value,
  formattedValue,
  col,
}: {
  value: unknown;
  formattedValue?: unknown;
  col?: BooleanColumnMeta;
}) {
  const s = useStyles();
  const checked = value === true || value === 1 || value === '1' || value === 'true';
  const label =
    (typeof formattedValue === 'string' && formattedValue) ||
    (checked ? (col?.trueOption?.label ?? 'Yes') : (col?.falseOption?.label ?? 'No'));
  return (
    <div className={s.booleanCell}>
      <Switch checked={checked} disabled />
      <Label size="small">{label}</Label>
    </div>
  );
}

/** Picklist / State / Status → Badge. Formatted value preferred for the label. */
function PicklistCellRenderer({
  value,
  formattedValue,
  col,
}: {
  value: unknown;
  formattedValue?: unknown;
  col?: PicklistColumnMeta | StatusColumnMeta | StateColumnMeta | MultiSelectPicklistColumnMeta;
}) {
  const s = useStyles();
  const isStateLike = col?.attributeType === 'State' || col?.attributeType === 'Status';
  const appearance = isStateLike ? 'filled' : 'tint';

  if (formattedValue != null && formattedValue !== '') {
    return (
      <div className={s.picklistCell}>
        <Badge appearance={appearance}>{String(formattedValue)}</Badge>
      </div>
    );
  }

  // Fallback: map raw integer through options[]
  const n = typeof value === 'number' ? value : parseInt(String(value), 10);
  if (Number.isNaN(n)) return <EmptyDash />;
  const opts = col?.options ?? [];
  const match = opts.find((o) => Number(o.value) === n);
  const label = match?.label ?? String(n);
  return (
    <div className={s.picklistCell}>
      <Badge appearance={appearance}>{label}</Badge>
    </div>
  );
}

/**
 * Lookup / Customer / Owner → bold name + small target-entity caption.
 * In OData the value column is `_<attr>_value` and the name lives in the
 * `@OData…FormattedValue` annotation; the target entity is in
 * `@Microsoft.Dynamics.CRM.lookuplogicalname`. Both passed in explicitly so
 * this renderer stays pure.
 */
function LookupCellRenderer({
  formattedValue,
  targetEntityDisplay,
}: {
  formattedValue?: unknown;
  /** Display name of the target entity, e.g. "Account" — small caption under the name */
  targetEntityDisplay?: string;
}) {
  const s = useStyles();
  if (formattedValue == null || formattedValue === '') return <EmptyDash />;
  return (
    <div className={s.lookupCell}>
      <span className={s.lookupName}>{String(formattedValue)}</span>
      {targetEntityDisplay && <span className={s.lookupType}>{targetEntityDisplay}</span>}
    </div>
  );
}

/** Numeric → right-aligned, tabular-nums. Formatted preferred (includes $, locale). */
type NumericLike =
  | MoneyColumnMeta
  | DecimalColumnMeta
  | DoubleColumnMeta
  | { attributeType: 'Integer'; precision?: undefined }
  | { attributeType: 'BigInt'; precision?: undefined };

function NumberCellRenderer({
  value,
  formattedValue,
  col,
}: {
  value: unknown;
  formattedValue?: unknown;
  col?: NumericLike;
}) {
  const s = useStyles();
  if (formattedValue != null && formattedValue !== '') {
    return <span className={s.numberCell}>{String(formattedValue)}</span>;
  }
  if (value == null || value === '') return <EmptyDash />;
  const n = typeof value === 'number' ? value : parseFloat(String(value));
  if (Number.isNaN(n)) return <EmptyDash />;

  const t = col?.attributeType;
  if (t === 'Money') {
    return (
      <span className={s.numberCell}>
        ${n.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
      </span>
    );
  }
  if (t === 'Integer' || t === 'BigInt') {
    return <span className={s.numberCell}>{n.toLocaleString()}</span>;
  }
  const precision = (col && 'precision' in col ? col.precision : undefined) ?? 2;
  return (
    <span className={s.numberCell}>
      {n.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: precision })}
    </span>
  );
}

/** DateTime → locale formatting, DateOnly vs DateAndTime aware. */
function DateTimeCellRenderer({
  value,
  formattedValue,
  col,
}: {
  value: unknown;
  formattedValue?: unknown;
  col?: DateTimeColumnMeta;
}) {
  const s = useStyles();
  if (formattedValue != null && formattedValue !== '') {
    return <span className={s.dateCell}>{String(formattedValue)}</span>;
  }
  if (!value) return <EmptyDash />;
  const d = new Date(String(value));
  if (Number.isNaN(d.getTime())) return <EmptyDash />;

  const isDateOnly = col?.format === 'DateOnly' || col?.dateTimeBehavior === 'DateOnly';
  if (isDateOnly) {
    return (
      <span className={s.dateCell}>
        {d.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' })}
      </span>
    );
  }
  return (
    <span className={s.dateCell}>
      {d.toLocaleString(undefined, {
        year: 'numeric',
        month: 'short',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
      })}
    </span>
  );
}

/** Default string fallback. Formatted preferred. Null/undefined → em-dash. */
function TextCellRenderer({ value, formattedValue }: { value: unknown; formattedValue?: unknown }) {
  const s = useStyles();
  const v = formattedValue ?? value;
  if (v == null || v === '') return <EmptyDash />;
  return <span className={s.truncateText}>{String(v)}</span>;
}

/**
 * Collection-valued expand cell — `contact_customer_accounts: [{…}, {…}]`.
 *
 * 1:N / N:N expands return an array of inner records. The flattener leaves
 * arrays in place (you can't fan a single account row across multiple
 * contact rows without changing grid semantics), so the renderer turns
 * the array into a count + brief preview pulled from the first 1–2
 * items. Users can drill into the JSON tab for the full nested shape.
 *
 * Preview rules:
 *   • Pick the first sensible "name-ish" field from each item:
 *     `fullname` → `name` → `title` → `subject` → otherwise nothing.
 *   • Show up to 2 items, then "+N more".
 *   • Empty array → em-dash.
 *   • If `@odata.nextLink` was present on the parent, surface "paginated".
 */
function CollectionCellRenderer({
  items,
  hasMore,
}: {
  items: Array<Record<string, unknown>>;
  hasMore?: boolean;
}) {
  const s = useStyles();
  if (!items.length) return <EmptyDash />;

  const NAME_KEYS = ['fullname', 'name', 'title', 'subject', 'emailaddress1'];
  const previewOf = (it: Record<string, unknown>): string | null => {
    for (const k of NAME_KEYS) {
      const v = it[k];
      if (typeof v === 'string' && v) return v;
    }
    return null;
  };

  const previews = items
    .slice(0, 2)
    .map(previewOf)
    .filter((v): v is string => !!v);
  const remaining = items.length - previews.length;

  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 6, minWidth: 0 }}>
      <Badge appearance="tint" color="informative">
        {items.length} {items.length === 1 ? 'item' : 'items'}
        {hasMore ? '+' : ''}
      </Badge>
      {previews.length > 0 && (
        <span
          className={s.truncateText}
          style={{ color: tokens.colorNeutralForeground2, fontSize: tokens.fontSizeBase200 }}
        >
          {previews.join(', ')}
          {remaining > 0 ? `, +${remaining} more` : ''}
        </span>
      )}
    </div>
  );
}

/**
 * Single entry point. Dispatches on column AttributeType. `targetEntityDisplay`
 * is only consulted by the Lookup renderer.
 */
export function getCellRenderer(args: {
  col: ColumnMeta | undefined;
  rawValue: unknown;
  formattedValue?: unknown;
  /** For lookup cells: human-readable target entity name */
  targetEntityDisplay?: string;
  /** True if the parent row had a `<key>@odata.nextLink` for this collection. */
  collectionHasMore?: boolean;
}) {
  const { col, rawValue, formattedValue, targetEntityDisplay, collectionHasMore } = args;

  if (rawValue == null && (formattedValue == null || formattedValue === '')) {
    return <EmptyDash />;
  }

  // Collection-valued expand — array of inner records. Detected BEFORE the
  // column-type dispatch since nav properties don't have a ColumnMeta and
  // would otherwise fall through to TextCellRenderer's `String([{…}])` =
  // "[object Object]" stringification.
  if (Array.isArray(rawValue) && rawValue.every((v) => v != null && typeof v === 'object')) {
    return (
      <CollectionCellRenderer
        items={rawValue as Array<Record<string, unknown>>}
        hasMore={collectionHasMore}
      />
    );
  }

  switch (col?.attributeType) {
    case 'Boolean':
      return <BooleanCellRenderer value={rawValue} formattedValue={formattedValue} col={col} />;
    case 'Picklist':
    case 'State':
    case 'Status':
    case 'MultiSelectPicklist':
      return (
        <PicklistCellRenderer
          value={rawValue}
          formattedValue={formattedValue}
          col={
            col as
              | PicklistColumnMeta
              | StatusColumnMeta
              | StateColumnMeta
              | MultiSelectPicklistColumnMeta
          }
        />
      );
    case 'Lookup':
    case 'Customer':
    case 'Owner':
      return (
        <LookupCellRenderer
          formattedValue={formattedValue}
          targetEntityDisplay={targetEntityDisplay}
        />
      );
    case 'Integer':
    case 'BigInt':
    case 'Decimal':
    case 'Double':
    case 'Money':
      return (
        <NumberCellRenderer
          value={rawValue}
          formattedValue={formattedValue}
          col={col as NumericLike}
        />
      );
    case 'DateTime':
      return <DateTimeCellRenderer value={rawValue} formattedValue={formattedValue} col={col} />;
    default:
      return <TextCellRenderer value={rawValue} formattedValue={formattedValue} />;
  }
}
