import { Field, SpinButton, Switch, Text, MessageBar, MessageBarBody, tokens, Caption1, ToggleButton } from '@fluentui/react-components';
import { NumberSymbol20Filled, Tag20Filled } from '@fluentui/react-icons';
import { PaneHead } from './PaneHead';
import { ApplyOverridesBanner } from './ApplyOverridesBanner';
import type { RequestGroup } from '../registry/requestTypes';

export function TopEditor({ top, setTop, maxPageSize, group = 'read' }: {
  top: number | null;
  setTop: (v: number | null) => void;
  /** Visible warning when both $top and Prefer:maxpagesize are set (mutually exclusive) */
  maxPageSize?: number | null;
  group?: RequestGroup;
}) {
  const conflict = top != null && top > 0 && maxPageSize != null && maxPageSize > 0;
  return (
    <div>
      <PaneHead
        icon={NumberSymbol20Filled}
        title="$top"
        sub="Maximum rows to return. Server hard cap is 5,000."
        group={group}
      />
      {conflict && (
        <MessageBar layout="multiline" intent="warning" style={{ marginBottom: 12 }}>
          <MessageBarBody>
            <strong>$top and Prefer: odata.maxpagesize are mutually exclusive</strong> — when both are set, <code>$top</code> is silently ignored.
          </MessageBarBody>
        </MessageBar>
      )}
      <div style={{ display: 'flex', flexDirection: 'column', gap: 14, maxWidth: 360 }}>
        <Field label="Top N" hint={top == null || top === 0 ? 'Empty / 0 = use server default (5,000).' : `Returns up to ${top} rows.`}>
          <SpinButton
            value={top ?? 0}
            min={0}
            max={5000}
            step={10}
            onChange={(_, d) => {
              const v = d.value ?? Number(d.displayValue ?? 0);
              setTop(v === 0 ? null : v);
            }}
          />
        </Field>
        <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
          {[5, 10, 25, 50, 100, 500, 1000, 5000].map(n => (
            <ToggleButton key={n} size="small" shape="circular" checked={top === n} onClick={() => setTop(n)}>
              {n.toLocaleString()}
            </ToggleButton>
          ))}
        </div>
        <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
          For large pages, prefer <code>Prefer: odata.maxpagesize</code> + <code>@odata.nextLink</code>.
        </Caption1>
      </div>
    </div>
  );
}

export function CountEditor({ countOn, setCountOn, group = 'read', applyActive }: {
  countOn: boolean;
  setCountOn: (b: boolean) => void;
  group?: RequestGroup;
  /** When true, $apply is on and $count is ignored. */
  applyActive?: boolean;
}) {
  return (
    <div>
      <PaneHead
        icon={Tag20Filled}
        title="$count"
        sub="Include the total count in the response (capped at 5,000 — see totalrecordcountlimitexceeded)."
        group={group}
      />
      {applyActive && <ApplyOverridesBanner clause="$count" />}
      <div style={{ display: 'flex', flexDirection: 'column', gap: 14, maxWidth: 480 }}>
        <Field label={countOn ? 'Including @odata.count' : 'Not including @odata.count'}>
          <Switch
            checked={countOn}
            onChange={(_, d) => setCountOn(d.checked)}
          />
        </Field>
        <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
          When you also need to know whether the count is capped, add the <code>Microsoft.Dynamics.CRM.totalrecordcount</code> and{' '}
          <code>totalrecordcountlimitexceeded</code> annotations via Prefer.
        </Text>
      </div>
    </div>
  );
}
