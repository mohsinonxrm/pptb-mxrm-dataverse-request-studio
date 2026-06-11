// NavPropertyPicker — Cardinality-aware navigation property picker.
//
// Cardinality is detected from metadata, not picked by the user. The UI
// reflects it: 1:N shows single related-record picker; N:N shows multi-pick.
//
// This picker lists every nav prop on the source table, surfaces its
// cardinality + target entity, and shows the implied HTTP verb so the user
// understands what they're about to send.
//
// Reference:
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/associate-disassociate-entities-using-web-api

import { Field, Combobox, Option, Badge, Caption1, tokens } from '@fluentui/react-components';
import { Link20Filled } from '@fluentui/react-icons';
import { PaneHead } from './PaneHead';
import { findTable, type NavProperty } from '../mock/metadata';
import { useLiveTable } from '../host/useLiveMetadata';
import type { RequestGroup } from '../registry/requestTypes';

export interface NavPropertyPickerProps {
  table: string;
  value: string | null;
  onChange: (navName: string | null) => void;
  /** Which group palette to use for the pane header. */
  group?: RequestGroup;
  /** Method context — Associate uses POST (collection) or PATCH (single-valued);
   *  Disassociate uses DELETE (collection) or PATCH (single-valued). */
  forOperation: 'associate' | 'disassociate';
  /** Optional banner shown at the bottom of the pane. */
  footer?: React.ReactNode;
}

export function NavPropertyPicker({
  table,
  value,
  onChange,
  group = 'relate',
  forOperation,
  footer,
}: NavPropertyPickerProps) {
  const tbl = findTable(table);
  const navProps = tbl?.navigationProperties ?? [];
  const current = navProps.find((n) => n.name === value);

  // Warm ONLY the currently-selected nav's target — the option rows can
  // safely render the raw logical name; loading 50+ targets up-front would
  // trip Dataverse's 100-concurrent-request cap.
  useLiveTable(current?.targetEntity ?? null);

  return (
    <div>
      <PaneHead
        icon={Link20Filled}
        title="Navigation property"
        sub="Cardinality is detected from metadata — the URL bar verb + shape adapt automatically."
        group={group}
      />

      <div style={{ maxWidth: 720 }}>
        <Field label="Navigation property">
          <Combobox
            value={current?.name ?? ''}
            selectedOptions={current ? [current.name] : []}
            onOptionSelect={(_, d) => onChange(d.optionValue ?? null)}
            placeholder="Pick a navigation property…"
          >
            {navProps.map((n) => (
              <Option key={n.name} value={n.name} text={n.name}>
                <NavOptionRow nav={n} forOperation={forOperation} />
              </Option>
            ))}
          </Combobox>
        </Field>

        {current && (
          <div style={{ marginTop: 16 }}>
            <NavSummaryCard nav={current} forOperation={forOperation} />
          </div>
        )}

        {footer && <div style={{ marginTop: 14 }}>{footer}</div>}
      </div>
    </div>
  );
}

// ──────────────────────────────────────────────────────────────
// Combobox option row — nav name + cardinality chip + target entity
// ──────────────────────────────────────────────────────────────
function NavOptionRow({
  nav,
  forOperation,
}: {
  nav: NavProperty;
  forOperation: 'associate' | 'disassociate';
}) {
  const target = findTable(nav.targetEntity);
  const verb = methodFor(nav, forOperation);
  return (
    <div style={{ display: 'flex', flexDirection: 'column', minWidth: 0 }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
        <span style={{ fontFamily: tokens.fontFamilyMonospace, fontWeight: 600 }}>{nav.name}</span>
        <Badge appearance="ghost" size="extra-small">
          {cardinalityShort(nav.cardinality)}
        </Badge>
        <Badge appearance="tint" color="brand" size="extra-small" style={{ marginLeft: 'auto' }}>
          {verb}
        </Badge>
      </div>
      <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
        → {target?.displayName ?? nav.targetEntity}
        {' · '}
        {valuedShort(nav.cardinality)}
      </Caption1>
    </div>
  );
}

// ──────────────────────────────────────────────────────────────
// Summary card — surfaces the resolved URL shape + method per cardinality
// ──────────────────────────────────────────────────────────────
function NavSummaryCard({
  nav,
  forOperation,
}: {
  nav: NavProperty;
  forOperation: 'associate' | 'disassociate';
}) {
  const target = findTable(nav.targetEntity);
  const verb = methodFor(nav, forOperation);
  const cardinalityKind = nav.cardinality === 'ManyToOne' ? 'single-valued' : 'collection-valued';
  // Verb chip color — matches the URL bar's method pill so the user can
  // visually link "this is what'll fire" between the two surfaces.
  const verbColor =
    verb === 'POST'
      ? tokens.colorPaletteGreenForeground1
      : verb === 'PATCH'
        ? tokens.colorPaletteDarkOrangeForeground1
        : tokens.colorPaletteRedForeground1; // DELETE

  return (
    <div
      style={{
        border: `1px solid ${tokens.colorNeutralStroke2}`,
        borderRadius: tokens.borderRadiusMedium,
        padding: 12,
        background: tokens.colorNeutralBackground1,
      }}
    >
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
        <Link20Filled style={{ width: 18, height: 18, color: tokens.colorBrandForeground1 }} />
        <strong style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 13 }}>{nav.name}</strong>
        <Badge appearance="tint" color="brand">
          {cardinalityShort(nav.cardinality)}
        </Badge>
        <span style={{ flexGrow: 1 }} />
        <Badge
          appearance="filled"
          style={{
            background: verbColor,
            color: tokens.colorNeutralForegroundOnBrand,
            fontWeight: 700,
          }}
        >
          {verb}
        </Badge>
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: '120px 1fr', rowGap: 4, fontSize: 11 }}>
        <span style={{ color: tokens.colorNeutralForeground3 }}>Target entity</span>
        <span style={{ fontFamily: tokens.fontFamilyMonospace }}>
          {target?.entitySetName ?? nav.targetEntity}{' '}
          <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
            · {target?.displayName ?? nav.targetEntity}
          </Caption1>
        </span>
        <span style={{ color: tokens.colorNeutralForeground3 }}>Cardinality</span>
        <span>
          {cardinalityKind} ({nav.cardinality})
        </span>
        <span style={{ color: tokens.colorNeutralForeground3 }}>Relationship</span>
        <span style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 10 }}>
          {nav.relationshipName}
        </span>
      </div>
    </div>
  );
}

function cardinalityShort(c: NavProperty['cardinality']): string {
  switch (c) {
    case 'OneToMany':
      return '1:N';
    case 'ManyToOne':
      return 'N:1';
    case 'ManyToMany':
      return 'N:N';
  }
}
function valuedShort(c: NavProperty['cardinality']): string {
  return c === 'ManyToOne' ? 'single-valued lookup' : 'collection-valued';
}
function methodFor(
  nav: NavProperty,
  op: 'associate' | 'disassociate',
): 'POST' | 'PATCH' | 'DELETE' {
  // Docs-preferred wire shape per cardinality:
  //   Associate    single-valued (N:1) → PATCH /<source>(<id>) with @odata.bind body
  //                collection-valued   → POST /<source>(<id>)/<nav>/$ref
  //   Disassociate single-valued       → PATCH /<source>(<id>) with @odata.bind: null
  //                collection-valued   → DELETE /<source>(<id>)/<nav>(<target>)/$ref
  // The PATCH branch is what makes single-valued Disassociate actually work in
  // PPTB — `dvHost.disassociate` requires a target id which single-valued doesn't
  // have, so the URL bar verb here was misleading prior to this rework.
  if (nav.cardinality === 'ManyToOne') return 'PATCH';
  return op === 'disassociate' ? 'DELETE' : 'POST';
}
