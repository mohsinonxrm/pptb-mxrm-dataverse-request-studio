// BypassEditor — single editor pane for all four MSCRM.* bypass headers.
//
// Shown as a sidebar clause node on every write mode. Replaces the
// inline-only "Bypass custom plug-ins" toggle that used to live on DeleteMode.
//
// References:
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/bypass-custom-business-logic
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/bypass-power-automate-flows
//
// Visual grammar: option cards inside a PaneHead-led section, with
// structured controls rather than a free-text header table.
// The actual header emission lives in engine/bypassHeaders.ts so the wire
// format is identical regardless of how this editor is laid out.

import { useState } from 'react';
import {
  Field,
  Radio,
  RadioGroup,
  Input,
  Switch,
  Caption1,
  Badge,
  Button,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  tokens,
} from '@fluentui/react-components';
import { Add20Regular, Delete20Regular, Warning20Filled } from '@fluentui/react-icons';
import { PaneHead } from './PaneHead';
import { useStudioStyles } from '../primitives/styles';
import type { BypassOptions, BypassBusinessLogicMode } from '../state/writeState';
import { STEP_IDS_DEFAULT_LIMIT, STEP_IDS_MAX_LIMIT } from '../engine/bypassHeaders';
import type { RequestGroup } from '../registry/requestTypes';

export interface BypassEditorProps {
  value: BypassOptions;
  onChange: (next: BypassOptions) => void;
  group?: RequestGroup;
}

export function BypassEditor({ value, onChange, group = 'write' }: BypassEditorProps) {
  const s = useStudioStyles();
  const [newStepId, setNewStepId] = useState('');

  const patch = (p: Partial<BypassOptions>) => onChange({ ...value, ...p });
  const setMode = (m: BypassBusinessLogicMode) => patch({ businessLogic: m });

  const stepCount = value.stepIds.filter(Boolean).length;
  const stepWarn = value.businessLogic === 'steps' && stepCount > STEP_IDS_DEFAULT_LIMIT;
  const stepErr = value.businessLogic === 'steps' && stepCount > STEP_IDS_MAX_LIMIT;

  const addStepId = () => {
    const v = newStepId.trim();
    if (!v) return;
    if (value.stepIds.includes(v)) {
      setNewStepId('');
      return;
    }
    patch({ stepIds: [...value.stepIds, v] });
    setNewStepId('');
  };
  const removeStepId = (i: number) => patch({ stepIds: value.stepIds.filter((_, j) => j !== i) });

  return (
    <div>
      <PaneHead
        icon={Warning20Filled}
        title="Bypass logic"
        sub="Skip server-side business logic or Power Automate triggers — privileged operations for bulk-data work."
        group={group}
      >
        <Badge appearance="ghost">{summarize(value)}</Badge>
      </PaneHead>

      <div style={{ display: 'flex', flexDirection: 'column', gap: 18, maxWidth: 720 }}>
        {/* ── Custom business logic ─────────────────────────────── */}
        <section>
          <SectionTitle>Custom business logic</SectionTitle>
          <Caption1
            style={{ color: tokens.colorNeutralForeground2, display: 'block', marginBottom: 8 }}
          >
            Emits <code>MSCRM.BypassBusinessLogicExecution</code> with the chosen scope. Requires
            the <code>prvBypassCustomBusinessLogic</code> privilege — by default only System
            Administrator has it.
          </Caption1>
          <RadioGroup
            value={value.businessLogic}
            onChange={(_, d) => setMode(d.value as BypassBusinessLogicMode)}
            layout="vertical"
          >
            <Radio
              value="none"
              label={<RadioLabel title="None" sub="Default — all custom logic runs." />}
            />
            <Radio
              value="sync"
              label={
                <RadioLabel
                  title="Sync only"
                  sub="Bypass synchronous plug-ins/workflows. Reduces per-record latency on bulk operations."
                />
              }
            />
            <Radio
              value="async"
              label={
                <RadioLabel
                  title="Async only"
                  sub="Bypass async logic (system jobs). Lets bulk operations complete without queueing thousands of async ops."
                />
              }
            />
            <Radio
              value="both"
              label={
                <RadioLabel
                  title="Both"
                  sub="Skip both sync and async logic. Most aggressive option."
                />
              }
            />
            <Radio
              value="steps"
              label={
                <RadioLabel
                  title="Specific steps"
                  sub="Bypass only the named plug-in step registrations (GUIDs). Useful when one plug-in causes the slowdown."
                />
              }
            />
          </RadioGroup>

          {value.businessLogic === 'sync' && (
            <div
              style={{
                marginTop: 10,
                padding: '8px 12px',
                borderRadius: tokens.borderRadiusMedium,
                background: tokens.colorNeutralBackground2,
              }}
            >
              <Switch
                checked={value.useLegacyHeader}
                onChange={(_, d) => patch({ useLegacyHeader: !!d.checked })}
                label={
                  <span style={{ fontSize: 12 }}>
                    Use legacy <code>MSCRM.BypassCustomPluginExecution</code> header
                    <Caption1 style={{ display: 'block', color: tokens.colorNeutralForeground3 }}>
                      Off: emit modern header. On: emit legacy (requires{' '}
                      <code>prvBypassCustomPlugins</code> privilege).
                    </Caption1>
                  </span>
                }
              />
            </div>
          )}

          {value.businessLogic === 'steps' && (
            <div className={s.inlineCard} style={{ padding: 12, marginTop: 10 }}>
              <strong style={{ fontSize: 12, display: 'block', marginBottom: 6 }}>
                Plug-in step IDs
              </strong>
              <Caption1
                style={{ color: tokens.colorNeutralForeground2, display: 'block', marginBottom: 8 }}
              >
                One GUID per row. Resolve them via{' '}
                <code>
                  GET
                  /sdkmessagefilters?$select=sdkmessagefilterid&amp;$filter=primaryobjecttypecode eq
                  '&lt;table&gt;' and sdkmessageid/name eq
                  '&lt;Op&gt;'&amp;$expand=sdkmessagefilterid_sdkmessageprocessingstep($select=name)
                </code>
                .
              </Caption1>
              <Field>
                <div style={{ display: 'flex', gap: 8 }}>
                  <Input
                    size="small"
                    value={newStepId}
                    onChange={(_, d) => setNewStepId(d.value)}
                    placeholder="00000000-0000-0000-0000-000000000000"
                    style={{ flex: 1 }}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') addStepId();
                    }}
                  />
                  <Button
                    icon={<Add20Regular />}
                    size="small"
                    onClick={addStepId}
                    disabled={!newStepId.trim()}
                  >
                    Add
                  </Button>
                </div>
              </Field>
              {value.stepIds.length > 0 && (
                <div style={{ marginTop: 8, display: 'flex', flexDirection: 'column', gap: 4 }}>
                  {value.stepIds.map((id, i) => (
                    <div
                      key={i}
                      style={{
                        display: 'flex',
                        alignItems: 'center',
                        gap: 8,
                        fontFamily: tokens.fontFamilyMonospace,
                        fontSize: 11,
                        padding: '4px 8px',
                        background: tokens.colorNeutralBackground1,
                        borderRadius: tokens.borderRadiusSmall,
                      }}
                    >
                      <span style={{ flex: 1 }}>{id}</span>
                      <Button
                        icon={<Delete20Regular />}
                        appearance="subtle"
                        size="small"
                        onClick={() => removeStepId(i)}
                      />
                    </div>
                  ))}
                </div>
              )}
              {stepErr && (
                <MessageBar layout="multiline" intent="error" style={{ marginTop: 8 }}>
                  <MessageBarBody>
                    <MessageBarTitle>Over the maximum step cap</MessageBarTitle>
                    Dataverse caps this list at {STEP_IDS_MAX_LIMIT}. The request will be rejected.
                  </MessageBarBody>
                </MessageBar>
              )}
              {!stepErr && stepWarn && (
                <MessageBar layout="multiline" intent="warning" style={{ marginTop: 8 }}>
                  <MessageBarBody>
                    Default org cap is {STEP_IDS_DEFAULT_LIMIT}. Anything over this requires the
                    <code> BypassBusinessLogicExecutionStepIdsLimit </code> OrgDbOrgSettings tweak
                    (max {STEP_IDS_MAX_LIMIT}).
                  </MessageBarBody>
                </MessageBar>
              )}
            </div>
          )}
        </section>

        {/* ── Power Automate flows ──────────────────────────────── */}
        <section>
          <SectionTitle>Power Automate flows</SectionTitle>
          <Caption1
            style={{ color: tokens.colorNeutralForeground2, display: 'block', marginBottom: 8 }}
          >
            Emits <code>MSCRM.SuppressCallbackRegistrationExpanderJob</code>. No privilege required,
            but flow owners are <strong>not notified</strong>
            that their logic was bypassed.
          </Caption1>
          <Switch
            checked={value.suppressFlows}
            onChange={(_, d) => patch({ suppressFlows: !!d.checked })}
            label="Suppress Power Automate flow triggers"
          />
          {value.suppressFlows && (
            <MessageBar layout="multiline" intent="warning" style={{ marginTop: 8 }}>
              <MessageBarBody>
                Use only when you've confirmed a backlog of{' '}
                <code>CallbackRegistration Expander Operation</code> jobs (operationtype 79 in
                <code> asyncoperation</code>). Communicate with flow owners before enabling for
                production traffic.
              </MessageBarBody>
            </MessageBar>
          )}
        </section>
      </div>
    </div>
  );
}

function SectionTitle({ children }: { children: React.ReactNode }) {
  return (
    <div
      style={{
        fontSize: 10,
        fontWeight: 700,
        textTransform: 'uppercase',
        letterSpacing: '0.5px',
        color: tokens.colorNeutralForeground3,
        marginBottom: 6,
      }}
    >
      {children}
    </div>
  );
}

function RadioLabel({ title, sub }: { title: string; sub: string }) {
  return (
    <span style={{ display: 'inline-flex', flexDirection: 'column' }}>
      <strong style={{ fontSize: 12 }}>{title}</strong>
      <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>{sub}</Caption1>
    </span>
  );
}

// Short summary for the sidebar / pane badge — e.g. "sync · flows".
export function summarize(b: BypassOptions): string {
  const parts: string[] = [];
  if (b.businessLogic === 'sync') parts.push(b.useLegacyHeader ? 'sync (legacy)' : 'sync');
  if (b.businessLogic === 'async') parts.push('async');
  if (b.businessLogic === 'both') parts.push('sync+async');
  if (b.businessLogic === 'steps') parts.push(`${b.stepIds.filter(Boolean).length} steps`);
  if (b.suppressFlows) parts.push('flows');
  return parts.length === 0 ? 'off' : parts.join(' · ');
}
