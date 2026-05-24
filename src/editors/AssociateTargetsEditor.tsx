// AssociateTargetsEditor — pick the related row(s) that will be linked to
// the source via the chosen navigation property.
//
// Layout:
//   • Card with avatars + names + GUIDs of selected targets
//   • Per-row dismiss button
//   • "+ Add target" picker (live RecordPicker over the target entity)
//   • Hint about $batch when multiple targets are queued
//   • "Generated requests" preview (handled by parent — passes the request list)
//
// Single-valued nav props: max 1 target. We clamp the UI to a single picker
// in that case so the user doesn't queue a request that the server would
// reject.
//
// Records come from the live `RecordPicker` typeahead — same primitive
// Merge / Delete / Update use; results fetched via `dataverseAPI.queryData`.

import { useState } from 'react';
import {
  Caption1, Badge, Tooltip, Button, Persona, MessageBar,
  MessageBarBody, MessageBarTitle, tokens, mergeClasses, Field,
} from '@fluentui/react-components';
import {
  Link20Filled, Dismiss20Regular, BoxMultiple20Regular, Add20Regular,
} from '@fluentui/react-icons';
import { useStudioStyles } from '../primitives/styles';
import { PaneHead } from './PaneHead';
import { findTable, type NavProperty } from '../mock/metadata';
import { RecordPicker } from '../primitives/RecordPicker';
import type { RequestGroup } from '../registry/requestTypes';

export interface AssociateTargetsEditorProps {
  /** Source table — used to look up the nav prop metadata. */
  table: string;
  /** Selected nav property name; drives the target entity + cardinality. */
  navProperty: string | null;
  /** Target row GUIDs the user has picked. */
  targets: string[];
  /** Replace the targets array. */
  setTargets: (next: string[]) => void;
  /**
   * Names captured when the user picked each target via the typeahead.
   * Persisted to the parent so the row cards keep showing names after
   * the picker resets, and so saved requests can rehydrate the labels.
   * Missing entries fall back to the GUID — RecordPicker.GuidPicker also
   * accepts hand-pasted GUIDs which won't have a name available.
   */
  targetNames?: Record<string, string>;
  setTargetNames?: (next: Record<string, string>) => void;
  group?: RequestGroup;
  /** Method that will be sent (POST or PUT) — surfaced as a chip. */
  resolvedMethod: 'POST' | 'PATCH' | null;
}

export function AssociateTargetsEditor({
  table, navProperty, targets, setTargets,
  targetNames: targetNamesProp, setTargetNames: setTargetNamesProp,
  group = 'relate', resolvedMethod,
}: AssociateTargetsEditorProps) {
  // Callers that don't track names externally get a transient session map.
  // Defined alongside the controlled path so the editor's display logic
  // doesn't need to branch.
  const [internalNames, setInternalNames] = useState<Record<string, string>>({});
  const targetNames = targetNamesProp ?? internalNames;
  const setTargetNames = setTargetNamesProp ?? setInternalNames;
  const s = useStudioStyles();
  const tbl = findTable(table);
  const nav = navProperty ? tbl?.navigationProperties.find(n => n.name === navProperty) : undefined;
  const targetTbl = nav ? findTable(nav.targetEntity) : undefined;
  const singleValued = nav?.cardinality === 'ManyToOne';
  const maxTargets = singleValued ? 1 : Infinity;

  // The "add target" picker resets to a clean state after every successful
  // pick. We tie its mounted identity to a counter so we can force a remount
  // (clears the typeahead text + radio mode) when the user adds a record.
  const [pickerKey, setPickerKey] = useState(0);

  if (!nav) {
    return (
      <div>
        <PaneHead icon={Link20Filled} title="Target records" sub="Pick a navigation property first." group={group} />
        <MessageBar layout="multiline" intent="info" style={{ maxWidth: 720 }}>
          <MessageBarBody>
            Switch to the <strong>Navigation property</strong> pane and pick one — the target entity comes from there.
          </MessageBarBody>
        </MessageBar>
      </div>
    );
  }

  const atCap = targets.length >= maxTargets;

  const onAdd = (id: string | null, name?: string) => {
    if (!id) return;
    if (targets.includes(id)) return; // dedupe — don't queue the same target twice
    if (atCap) return;
    const nextTargets = singleValued ? [id] : [...targets, id];
    setTargets(nextTargets);
    if (name) setTargetNames({ ...targetNames, [id]: name });
    setPickerKey(k => k + 1); // remount the picker so the typeahead clears
  };

  const onRemove = (id: string) => {
    setTargets(targets.filter(t => t !== id));
    if (targetNames[id] != null) {
      const next = { ...targetNames };
      delete next[id];
      setTargetNames(next);
    }
  };

  return (
    <div>
      <PaneHead
        icon={Link20Filled}
        title={singleValued ? 'Target record' : 'Target records'}
        sub={
          <span>
            {singleValued
              ? <>One target is allowed for <code>{nav.cardinality}</code> nav props — the server PUT sets the lookup.</>
              : <>One or more targets — each fires a separate POST request. For many targets, Dataverse recommends a <code>$batch</code> request.</>}
          </span>
        }
        group={group}
      >
        <Badge appearance="tint" color={targets.length > 0 ? 'brand' : 'subtle'}>
          {targets.length} target{targets.length === 1 ? '' : 's'}
        </Badge>
        {resolvedMethod && (
          <Badge appearance="filled" color={resolvedMethod === 'POST' ? 'success' : 'informative'} style={{ fontWeight: 700 }}>
            {resolvedMethod}
          </Badge>
        )}
        {!singleValued && targets.length > 1 && (
          <Tooltip content="Will fire one request per target — Dataverse recommends $batch for large counts. Batch transport lands in a later pass." relationship="description">
            <Badge appearance="ghost" icon={<BoxMultiple20Regular style={{ width: 12, height: 12 }} />}>
              {targets.length} requests queued
            </Badge>
          </Tooltip>
        )}
      </PaneHead>

      {/* Selected targets card */}
      <div className={mergeClasses(s.inlineCard)} style={{ padding: 12, marginBottom: 14, maxWidth: 880 }}>
        {targets.length === 0 ? (
          <Caption1 style={{ padding: '12px 4px', color: tokens.colorNeutralForeground3, fontStyle: 'italic', display: 'block', textAlign: 'center' }}>
            No targets selected — use the picker below to choose {singleValued ? 'one' : 'one or more'} <strong>{targetTbl?.displayName ?? nav.targetEntity}</strong> record{singleValued ? '' : 's'}.
          </Caption1>
        ) : (
          <div style={{ display: 'flex', flexDirection: 'column' }}>
            {targets.map((id, idx) => {
              const name = targetNames[id];
              return (
                <div
                  key={id}
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    gap: 10,
                    padding: '8px 4px',
                    borderBottom: idx < targets.length - 1 ? `1px solid ${tokens.colorNeutralStroke3}` : 'none',
                  }}
                >
                  <Persona size="small" name={name ?? '?'} avatar={{ color: 'colorful' }} />
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontSize: 12, fontWeight: 600 }}>
                      {name ?? <span style={{ fontStyle: 'italic', color: tokens.colorNeutralForeground3 }}>(name not resolved)</span>}
                    </div>
                    <div style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 10, color: tokens.colorNeutralForeground3 }}>
                      /{targetTbl?.entitySetName ?? '?'}({id})
                    </div>
                  </div>
                  <Tooltip content="Remove target" relationship="label">
                    <Button
                      size="small"
                      appearance="subtle"
                      icon={<Dismiss20Regular />}
                      onClick={() => onRemove(id)}
                      aria-label="Remove target"
                    />
                  </Tooltip>
                </div>
              );
            })}
          </div>
        )}
      </div>

      {/* Live add picker — RecordPicker with typeahead + by-GUID toggle.
          Keyed on `pickerKey` so adding a target wipes the input cleanly. */}
      {!atCap && (
        <div style={{ maxWidth: 560 }}>
          <Field
            label={singleValued ? 'Pick target' : 'Add target'}
            hint={
              singleValued
                ? `Pick a single ${targetTbl?.displayName ?? 'record'} — the PUT sets this as the lookup value.`
                : `Each pick queues a separate POST. Currently ${targets.length} target${targets.length === 1 ? '' : 's'} queued.`
            }
          >
            <RecordPicker
              key={pickerKey}
              table={targetTbl?.logicalName ?? nav.targetEntity}
              selectedId={null}
              onPick={(r) => onAdd(r?.id ?? null, r?.primary)}
              placeholder={`Search ${targetTbl?.displayName ?? 'record'}s…`}
            />
          </Field>
          {!singleValued && (
            <Caption1 style={{ display: 'block', marginTop: 6, color: tokens.colorNeutralForeground3 }}>
              <Add20Regular style={{ width: 12, height: 12, verticalAlign: 'middle' }} /> picking a record adds it to the queue and resets the picker for the next pick.
            </Caption1>
          )}
        </div>
      )}

      {atCap && singleValued && (
        <MessageBar layout="multiline" intent="info" style={{ maxWidth: 720, marginTop: 12 }}>
          <MessageBarBody>
            <MessageBarTitle>Single-valued nav property — one target only.</MessageBarTitle>
            To change the target, remove the current one first. The server PUT atomically replaces the previous lookup value.
          </MessageBarBody>
        </MessageBar>
      )}
    </div>
  );
}

// Suppress unused-import warning
export const _AssociateTargetsEditor_nav: (n: NavProperty) => void = () => undefined;
