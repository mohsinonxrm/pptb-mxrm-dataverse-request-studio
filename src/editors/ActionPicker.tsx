// ActionPicker — searchable Combobox of OData actions / functions from the
// mock CSDL. The picker filters by category (OOB / Custom API / Custom Action /
// Function) so the same component drives every Execute-* mode.
//
// Each option surfaces:
//   • the action name (mono),
//   • bound vs unbound + bound entity,
//   • a short description from the CSDL.
//
// Reference:
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/use-web-api-actions
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/use-web-api-functions

import { useEffect, useMemo, useState } from 'react';
import { Field, Combobox, Option, Badge, Caption1, tokens, Spinner, MessageBar, MessageBarBody } from '@fluentui/react-components';
import { Flash20Filled, Code20Filled, BoxMultiple20Regular } from '@fluentui/react-icons';
import { PaneHead } from './PaneHead';
import { type CsdlAction } from '../mock/actionsCsdl';
import { findTable } from '../mock/metadata';
import { actions as actionsProvider } from '../host/csdlProvider';
import type { RequestGroup } from '../registry/requestTypes';

export interface ActionPickerProps {
  /** Filter the CSDL — drives which subset of actions/functions is offered. */
  filter: (a: CsdlAction) => boolean;
  /** Currently selected action name (or null). */
  value: string | null;
  /** Display title for the pane head. */
  title: string;
  /** Inline subtitle / context line. */
  sub?: string;
  /** Called with the new selection (or null when cleared). */
  onChange: (name: string | null) => void;
  group?: RequestGroup;
}

// Scope sentinel: empty string = no scope picked yet (operation dropdown
// stays disabled). `__unbound__` = the "Unbound" entry. Any other string
// is a binding entity's logical name. Per design: there is NO "Any scope"
// option — scope must be either Unbound or a specific entity to surface
// operations. Clearing the scope clears the operation too (cascade reset).
const SCOPE_UNBOUND = '__unbound__';
const SCOPE_EMPTY = '';

export function ActionPicker({
  filter, value, title, sub, onChange, group = 'execute',
}: ActionPickerProps) {
  // Two-dropdown UX. Each has TWO state fields:
  //   • The committed value (`scope` / `value` — the latter is parent-owned)
  //   • The typed-input text (`scopeInput` / `opInput`) for freeform filtering
  // Splitting them prevents the bleed bug where typed filter text from a
  // prior interaction kept showing as the "current selection".
  const [scope, setScope] = useState<string>(SCOPE_EMPTY);
  const [scopeInput, setScopeInput] = useState<string>('');
  const [opInput, setOpInput] = useState<string>('');

  // Live CSDL load — actionsProvider routes to dvHost.metadata.getCSDLDocument()
  // in PPTB (cached for the session) and falls back to the mock catalog in
  // standalone. We load once on mount; the provider's internal cache makes
  // repeat mounts free. The CSDL XML is 1-5 MB → first load may take 1-2s.
  const [loaded, setLoaded] = useState<CsdlAction[] | null>(null);
  const [loadError, setLoadError] = useState<string | null>(null);
  useEffect(() => {
    let cancelled = false;
    actionsProvider.loadAll()
      .then(list => { if (!cancelled) setLoaded(list); })
      .catch(e => { if (!cancelled) setLoadError(e instanceof Error ? e.message : String(e)); });
    return () => { cancelled = true; };
  }, []);

  const loading = loaded === null && !loadError;
  const all = useMemo(() => (loaded ?? []).filter(filter), [loaded, filter]);

  // ── Scope options (distinct binding entities) ──
  // Walks every op in `all` collecting unbound count + per-entity counts.
  const scopeOptions = useMemo(() => {
    let unboundCount = 0;
    const entityCounts = new Map<string, number>();
    for (const a of all) {
      if (a.binding.kind === 'unbound') {
        unboundCount++;
      } else {
        const e = a.binding.entityType;
        entityCounts.set(e, (entityCounts.get(e) ?? 0) + 1);
      }
    }
    const entities = Array.from(entityCounts.entries())
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([logicalName, count]) => ({ logicalName, count }));
    return { unboundCount, entities };
  }, [all]);

  // Filter scope-dropdown options by typed input. Lets the user actually
  // narrow a long entity list (e.g. ~700 entities in a customized org)
  // by typing — the previous build ignored typed text.
  const scopeQuery = scopeInput.trim().toLowerCase();
  // Only treat the input as a filter if it doesn't already match the
  // committed selection's display text (avoids "filtered to one match"
  // confusion right after a selection lands).
  const committedDisplay = scope === SCOPE_EMPTY ? '' : scope === SCOPE_UNBOUND ? 'Unbound' : scope;
  const isFilteringScope = scopeQuery.length > 0 && scopeQuery !== committedDisplay.toLowerCase();
  const filteredEntities = useMemo(() => {
    if (!isFilteringScope) return scopeOptions.entities;
    const q = scopeQuery;
    return scopeOptions.entities.filter(s => {
      if (s.logicalName.toLowerCase().includes(q)) return true;
      const tt = findTable(s.logicalName);
      return tt ? tt.displayName.toLowerCase().includes(q) : false;
    });
  }, [scopeOptions.entities, isFilteringScope, scopeQuery]);
  const showUnboundOption =
    scopeOptions.unboundCount > 0 &&
    (!isFilteringScope || 'unbound'.includes(scopeQuery));

  // ── Operation list filtered by scope + search ──
  //
  // No scope picked → no operations shown. The picker requires the user to
  // pick scope first (Unbound or a specific entity); only then does the
  // Operation dropdown populate. This cascading-required model removes the
  // ambiguity of an "Any scope" option that defeated the purpose of having
  // scope at all.
  const inScope = useMemo(() => {
    if (scope === SCOPE_EMPTY) return [] as CsdlAction[];
    if (scope === SCOPE_UNBOUND) return all.filter(a => a.binding.kind === 'unbound');
    return all.filter(a => a.binding.kind !== 'unbound' && a.binding.entityType === scope);
  }, [all, scope]);

  // Op-input is freeform — treat it as a filter unless it already matches the
  // committed selection's display name (post-selection state).
  const current = value ? all.find(a => a.name === value) : undefined;
  const committedOpDisplay = current?.displayName ?? current?.name ?? '';
  const isFilteringOp = opInput.trim().length > 0 && opInput.trim() !== committedOpDisplay;
  const filtered = useMemo(() => {
    if (!isFilteringOp) return inScope;
    const q = opInput.toLowerCase();
    return inScope.filter(a =>
      a.name.toLowerCase().includes(q) ||
      (a.displayName?.toLowerCase().includes(q) ?? false) ||
      (a.description?.toLowerCase().includes(q) ?? false));
  }, [inScope, opInput, isFilteringOp]);

  // Sync the op-input text to the selected operation's display name when
  // the selection changes externally (mode switch, Save load, scope-
  // induced clear) AND when the CSDL data lands (`loaded` transition).
  //
  // The `loaded` dependency is critical for the remount case:
  //   • On remount, `value='QualifyLead'` (preserved by parent) but
  //     `loaded` starts null. `all=[]`, `current=undefined`,
  //     `committedOpDisplay=''` → opInput would stay empty.
  //   • Once `loaded` populates, `current` resolves and the effect
  //     re-fires → opInput picks up the correct display name.
  //
  // Without it, the parent's persisted value silently fails to surface
  // in the dropdown after navigating between panes.
  useEffect(() => {
    setOpInput(committedOpDisplay);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [value, loaded]);

  // Sync the scope-input text to the committed scope display when it
  // changes externally. Same reasoning as above.
  useEffect(() => {
    setScopeInput(committedDisplay);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [scope]);

  // ── Recover scope from value on (re)mount / data load ──
  //
  // `scope` is local component state — it resets to SCOPE_EMPTY when the
  // ActionPicker remounts (e.g. user navigates from the picker pane to
  // the params pane and back). The parent's `value` (operation name)
  // however persists across navigation. Without this effect, you get:
  //
  //   • value = 'QualifyLead' (preserved)
  //   • scope = ''            (lost on remount)
  //   • → dropdowns appear empty, BUT the selected-operation card still
  //     renders because `current = all.find(a => a.name === value)` resolves.
  //
  // Here we derive the right scope from the value's binding the moment
  // both are available (value set + CSDL loaded). One-shot recovery:
  // we only set scope when it's currently EMPTY — never overwrite an
  // explicit user choice.
  useEffect(() => {
    if (!value) return;
    if (scope !== SCOPE_EMPTY) return;
    if (!loaded) return;
    const action = loaded.find(a => a.name === value);
    if (!action) return;
    const derived = action.binding.kind === 'unbound'
      ? SCOPE_UNBOUND
      : action.binding.entityType;
    setScope(derived);
  }, [value, loaded, scope]);

  // ── Cascading reset ──
  //
  // Scope is the parent, operation is the child. Any change to scope
  // that makes the current op invalid (scope cleared OR scope changed to
  // a different value) drops the op selection so the user re-picks from
  // the new (or empty) list.
  useEffect(() => {
    if (!current) return;
    // Scope cleared → drop op. The op dropdown will be disabled until the
    // user picks a scope again.
    if (scope === SCOPE_EMPTY) {
      onChange(null);
      setOpInput('');
      return;
    }
    const matches =
      (scope === SCOPE_UNBOUND && current.binding.kind === 'unbound') ||
      (scope !== SCOPE_UNBOUND && current.binding.kind !== 'unbound' && current.binding.entityType === scope);
    if (!matches) {
      onChange(null);
      setOpInput('');
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [scope]);

  return (
    <div>
      <PaneHead
        icon={Flash20Filled}
        title={title}
        sub={sub ?? 'Browse the CSDL — pick an operation to populate its typed parameter form.'}
        group={group}
      >
        {/* Per design feedback — the Combobox below is freeform + clearable,
            so the user can filter by typing into it directly. The separate
            "Filter…" input that lived here was redundant and conflicted
            with the combobox's own search state. */}
        {loading
          ? <Badge appearance="ghost" icon={<Spinner size="extra-tiny" />}>loading CSDL…</Badge>
          : <Badge appearance="ghost">{all.length} available</Badge>}
      </PaneHead>

      {loadError && (
        <MessageBar intent="error" layout="multiline" style={{ marginBottom: 12, maxWidth: 760 }}>
          <MessageBarBody>
            Couldn't load CSDL metadata: {loadError}
          </MessageBarBody>
        </MessageBar>
      )}

      <div style={{ maxWidth: 760, display: 'flex', flexDirection: 'column', gap: 14 }}>
        {/* ── Scope dropdown — narrows the operation list ────────────── */}
        {/*
          Cascading-required filter: scope MUST be picked (Unbound or a
          specific entity) before the Operation dropdown enables. Each
          scope row shows the count of operations under it so the user
          can see at-a-glance which entities have coverage.
        */}
        <Field label="Scope" hint="Pick Unbound or a binding entity. Required before choosing an operation.">
          <Combobox
            freeform
            clearable
            value={scopeInput}
            selectedOptions={scope === SCOPE_EMPTY ? [] : [scope]}
            onChange={(e) => {
              const next = (e.target as HTMLInputElement).value;
              setScopeInput(next);
              // Empty input (or X clear) → reset to no-scope. The cascade
              // effect upstream will clear the operation too.
              if (next === '' && scope !== SCOPE_EMPTY) setScope(SCOPE_EMPTY);
            }}
            onOptionSelect={(_, d) => {
              const next = d.optionValue ?? SCOPE_EMPTY;
              setScope(next);
              // Sync the input text immediately so the dropdown closes with
              // the picked value displayed (not the typed filter text).
              const display =
                next === SCOPE_EMPTY ? '' :
                next === SCOPE_UNBOUND ? 'Unbound' :
                next;
              setScopeInput(display);
            }}
            placeholder={
              loading ? 'Loading…' :
              all.length === 0 ? 'No operations available' :
              `Pick scope · ${scopeOptions.unboundCount} unbound, ${scopeOptions.entities.length} entit${scopeOptions.entities.length === 1 ? 'y' : 'ies'}`
            }
            disabled={loading || all.length === 0}
            listbox={{ style: { maxHeight: 360 } }}
            style={{ width: '100%' }}
          >
            {showUnboundOption && (
              <Option value={SCOPE_UNBOUND} text="Unbound">
                <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                  <Badge appearance="ghost" size="extra-small">Unbound</Badge>
                  <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                    {scopeOptions.unboundCount} op{scopeOptions.unboundCount === 1 ? '' : 's'} · invoked without a source record
                  </Caption1>
                </div>
              </Option>
            )}
            {filteredEntities.map(s => {
              const tt = findTable(s.logicalName);
              return (
                <Option key={s.logicalName} value={s.logicalName} text={s.logicalName}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 6, minWidth: 0 }}>
                    <span style={{ fontFamily: tokens.fontFamilyMonospace, fontWeight: 600 }}>{s.logicalName}</span>
                    {tt && (
                      <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                        {tt.displayName}
                      </Caption1>
                    )}
                    <span style={{ flexGrow: 1 }} />
                    <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>{s.count} op{s.count === 1 ? '' : 's'}</Caption1>
                  </div>
                </Option>
              );
            })}
            {filteredEntities.length === 0 && !showUnboundOption && (
              <Option value="__noscope" text="" disabled>
                <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                  No scopes match "{scopeQuery}"
                </Caption1>
              </Option>
            )}
          </Combobox>
        </Field>

        <Field label="Operation">
          <Combobox
            // Filterable + clearable. opInput is the typed text (separate
            // from `value`, the committed selection). Effects above sync
            // opInput to the selection's display name when value changes
            // externally (mode switch / Save load / scope-induced clear).
            freeform
            clearable
            value={opInput}
            selectedOptions={value ? [value] : []}
            onChange={(e) => {
              const next = (e.target as HTMLInputElement).value;
              setOpInput(next);
              // Empty input also clears the committed selection.
              if (next === '' && value) onChange(null);
            }}
            onOptionSelect={(_, d) => {
              const next = d.optionValue ?? null;
              onChange(next);
              if (next) {
                const picked = all.find(a => a.name === next);
                setOpInput(picked?.displayName ?? picked?.name ?? '');
              } else {
                setOpInput('');
              }
            }}
            disabled={loading || scope === SCOPE_EMPTY}
            placeholder={
              loading ? 'Loading actions from $metadata…' :
              loadError ? 'CSDL load failed — see error above' :
              scope === SCOPE_EMPTY ? 'Pick a scope above first' :
              inScope.length === 0 ? `No ${all[0]?.kind?.toLowerCase() ?? 'operation'}s in this scope` :
              `Search ${inScope.length} ${all[0]?.kind ?? 'item'}${inScope.length === 1 ? '' : 's'}…`
            }
            listbox={{ style: { maxHeight: 420 } }}
            style={{ width: '100%' }}
          >
            {filtered.map(a => (
              <Option key={a.name} value={a.name} text={a.displayName ?? a.name}>
                <ActionOptionRow action={a} />
              </Option>
            ))}
            {filtered.length === 0 && (
              <Option value="__none" text="" disabled>
                <span style={{ color: tokens.colorNeutralForeground3, fontSize: 12 }}>No matches</span>
              </Option>
            )}
          </Combobox>
        </Field>

        {current && (
          <div style={{ marginTop: 16 }}>
            <ActionSummaryCard action={current} />
          </div>
        )}
      </div>
    </div>
  );
}

function ActionOptionRow({ action }: { action: CsdlAction }) {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', minWidth: 0 }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
        <span style={{ fontFamily: tokens.fontFamilyMonospace, fontWeight: 600 }}>{action.name}</span>
        <Badge appearance="ghost" size="extra-small">{action.kind}</Badge>
        {action.isComposable && <Badge appearance="ghost" size="extra-small">composable</Badge>}
        <Badge appearance="tint" color="brand" size="extra-small" style={{ marginLeft: 'auto' }}>
          {action.kind === 'Action' ? 'POST' : 'GET'}
        </Badge>
      </div>
      <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
        {bindingDescription(action)}
        {action.description && <> · {action.description.slice(0, 110)}{action.description.length > 110 ? '…' : ''}</>}
      </Caption1>
    </div>
  );
}

function ActionSummaryCard({ action }: { action: CsdlAction }) {
  const boundTbl = action.binding.kind !== 'unbound' ? findTable(action.binding.entityType) : undefined;
  return (
    <div style={{
      border: `1px solid ${tokens.colorNeutralStroke2}`,
      borderRadius: tokens.borderRadiusMedium,
      padding: 12,
      background: tokens.colorNeutralBackground1,
    }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
        {action.kind === 'Function' ? <Code20Filled style={{ color: tokens.colorBrandForeground1 }} /> : <Flash20Filled style={{ color: tokens.colorBrandForeground1 }} />}
        <strong style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 13 }}>{action.name}</strong>
        <Badge appearance="tint" color="brand">{action.kind}</Badge>
        {action.isComposable && <Badge appearance="ghost" icon={<BoxMultiple20Regular style={{ width: 10, height: 10 }} />}>composable</Badge>}
        <span style={{ flexGrow: 1 }} />
        <Badge appearance="filled" color={action.kind === 'Action' ? 'success' : 'brand'} style={{ fontWeight: 700 }}>
          {action.kind === 'Action' ? 'POST' : 'GET'}
        </Badge>
      </div>
      {action.description && (
        <Caption1 style={{ color: tokens.colorNeutralForeground2, display: 'block', marginBottom: 8 }}>
          {action.description}
        </Caption1>
      )}
      <div style={{ display: 'grid', gridTemplateColumns: '120px 1fr', rowGap: 4, fontSize: 11 }}>
        <span style={{ color: tokens.colorNeutralForeground3 }}>Namespace</span>
        <span style={{ fontFamily: tokens.fontFamilyMonospace }}>{action.namespace}</span>
        <span style={{ color: tokens.colorNeutralForeground3 }}>Binding</span>
        <span>{bindingDescription(action)}{boundTbl && <> · <span style={{ fontFamily: tokens.fontFamilyMonospace, color: tokens.colorNeutralForeground3 }}>{boundTbl.entitySetName}</span></>}</span>
        <span style={{ color: tokens.colorNeutralForeground3 }}>Parameters</span>
        <span>
          {action.parameters.length === 0
            ? <em style={{ color: tokens.colorNeutralForeground3 }}>(none)</em>
            : action.parameters.map(p => (
                <Badge key={p.name} appearance="ghost" size="extra-small" style={{ marginRight: 4, fontFamily: tokens.fontFamilyMonospace }}>
                  {p.name}: {p.type}{p.required ? ' *' : ''}
                </Badge>
              ))}
        </span>
        <span style={{ color: tokens.colorNeutralForeground3 }}>Return type</span>
        <span style={{ fontFamily: tokens.fontFamilyMonospace }}>
          {action.returnType.typeName}
          <Caption1 style={{ marginLeft: 6, color: tokens.colorNeutralForeground3, fontFamily: tokens.fontFamilyBase }}>
            ({action.returnType.kind})
          </Caption1>
        </span>
        <span style={{ color: tokens.colorNeutralForeground3 }}>Source</span>
        <span><Badge appearance="ghost">{action.source}</Badge></span>
      </div>
    </div>
  );
}

function bindingDescription(action: CsdlAction): string {
  switch (action.binding.kind) {
    case 'unbound':    return 'Unbound';
    case 'entity':     return `Bound to ${action.binding.entityType}`;
    case 'collection': return `Bound to ${action.binding.entityType} collection`;
  }
}
