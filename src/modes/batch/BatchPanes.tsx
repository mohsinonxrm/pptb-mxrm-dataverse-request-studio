// Overview + group panes for the $batch composer (#9).

import {
  Button,
  Caption1,
  Menu,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  Switch,
  ToggleButton,
  tokens,
} from '@fluentui/react-components';
import {
  Add20Regular,
  ArrowDown20Regular,
  ArrowUp20Regular,
  Delete20Regular,
  History20Regular,
  LockClosed20Regular,
  Stack20Regular,
  TextBulletListSquare20Regular,
} from '@fluentui/react-icons';
import { PaneHead } from '../../editors/PaneHead';
import { SegmentedToggle } from '../../primitives/SegmentedToggle';
import { MethodPill } from '../../primitives/MethodPill';
import {
  itemSummary,
  type BatchGroup,
  type BatchOpKind,
  type DispatchSegment,
} from '../../engine/batch/batchModel';

export const OP_MENU: { id: BatchOpKind; label: string }[] = [
  { id: 'create', label: 'Create' },
  { id: 'update', label: 'Update' },
  { id: 'delete', label: 'Delete' },
  { id: 'retrieve', label: 'Retrieve' },
  { id: 'associate', label: 'Associate' },
  { id: 'custom', label: 'Custom request' },
];

export function AddRequestMenu({
  onAdd,
  label = 'Add request',
  disabledGet,
}: {
  onAdd: (op: BatchOpKind) => void;
  label?: string;
  /** Change sets can't hold GETs — hide Retrieve there. */
  disabledGet?: boolean;
}) {
  return (
    <Menu>
      <MenuTrigger disableButtonEnhancement>
        <Button icon={<Add20Regular />}>{label}</Button>
      </MenuTrigger>
      <MenuPopover>
        <MenuList>
          {OP_MENU.filter((o) => !(disabledGet && o.id === 'retrieve')).map((o) => (
            <MenuItem key={o.id} onClick={() => onAdd(o.id)}>
              {o.label}
            </MenuItem>
          ))}
        </MenuList>
      </MenuPopover>
    </Menu>
  );
}

export function BatchOverviewPane({
  groups,
  continueOnError,
  setContinueOnError,
  plan,
  unavailable,
  needsConsent,
  orgUrl,
  onAddGroup,
  onOpenSaved,
  onAddLast,
  lastLabel,
  onSelect,
  notice,
  variant = 'batch',
  returnResponses = true,
  setReturnResponses,
}: {
  variant?: 'batch' | 'execute-multiple';
  returnResponses?: boolean;
  setReturnResponses?: (v: boolean) => void;
  groups: BatchGroup[];
  continueOnError: boolean;
  setContinueOnError: (v: boolean) => void;
  plan: DispatchSegment[];
  unavailable: string | null;
  needsConsent: boolean;
  orgUrl: string;
  onAddGroup: (atomic: boolean) => void;
  onOpenSaved: () => void;
  onAddLast: () => void;
  lastLabel: string | null;
  onSelect: (id: string) => void;
  notice: string | null;
}) {
  const total = groups.reduce((n, g) => n + g.items.length, 0);
  const em = variant === 'execute-multiple';
  let n = 0;
  return (
    <div style={{ maxWidth: 860 }}>
      <PaneHead
        icon={Stack20Regular}
        title={em ? 'ExecuteMultiple' : '$batch request'}
        sub={
          em
            ? 'Up to 1,000 independent requests in one call, with the SDK’s ContinueOnError and ReturnResponses settings.'
            : 'Up to 1,000 requests in one POST to /$batch. Requests in a change set succeed or roll back together.'
        }
        group="batch"
      />

      {em && (
        <MessageBar intent="info" layout="multiline" style={{ marginBottom: 12 }}>
          <MessageBarBody>
            <MessageBarTitle>Sent as a Web API $batch</MessageBarTitle>
            ExecuteMultiple is an SDK for .NET (Organization service) message; the Web API has no
            ExecuteMultiple action. Its Web API equivalent is a <code>$batch</code> of independent
            requests, which is what DRS sends. ContinueOnError maps to{' '}
            <code>Prefer: odata.continue-on-error</code>; ReturnResponses = false shows only the
            faults, as ExecuteMultipleResponse does. The Code tab has the ExecuteMultipleRequest
            (C#). For all-or-nothing, use the $batch Request mode’s change sets
            (ExecuteTransaction).
          </MessageBarBody>
        </MessageBar>
      )}

      {unavailable && (
        <MessageBar intent="warning" layout="multiline" style={{ marginBottom: 12 }}>
          <MessageBarBody>
            <MessageBarTitle>Can’t send from here</MessageBarTitle>
            {unavailable} You can still compose the batch and copy it from the Code tab.
          </MessageBarBody>
        </MessageBar>
      )}
      {notice && (
        <MessageBar intent="info" layout="multiline" style={{ marginBottom: 12 }}>
          <MessageBarBody>{notice}</MessageBarBody>
        </MessageBar>
      )}

      <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginBottom: 16 }}>
        {em ? (
          groups[0] && (
            <Button icon={<Add20Regular />} onClick={() => onSelect(`group:${groups[0].id}`)}>
              Add request
            </Button>
          )
        ) : (
          <>
            <Button icon={<LockClosed20Regular />} onClick={() => onAddGroup(true)}>
              Add change set
            </Button>
            <Button icon={<TextBulletListSquare20Regular />} onClick={() => onAddGroup(false)}>
              Add independent group
            </Button>
          </>
        )}
        <Button icon={<Add20Regular />} onClick={onOpenSaved}>
          Add from Saved…
        </Button>
        <Button icon={<History20Regular />} onClick={onAddLast} disabled={!lastLabel}>
          {lastLabel ? `Add last request (${lastLabel})` : 'Add last request'}
        </Button>
      </div>

      <div
        style={{
          display: 'flex',
          gap: 12,
          alignItems: 'flex-start',
          padding: 12,
          border: `1px solid ${tokens.colorNeutralStroke2}`,
          borderRadius: 6,
          background: tokens.colorNeutralBackground1,
          marginBottom: 16,
        }}
      >
        <Switch
          checked={continueOnError}
          onChange={(_, d) => setContinueOnError(d.checked)}
          aria-label="Continue on error"
        />
        <div>
          <strong>{em ? 'ContinueOnError' : 'Continue on error'}</strong>{' '}
          <code style={{ fontSize: 12 }}>Prefer: odata.continue-on-error</code>
          <Caption1 block style={{ color: tokens.colorNeutralForeground3 }}>
            {em
              ? 'False: processing stops at the first fault; the rest don’t run. True: every request runs and each fault is reported. Requests that already ran aren’t undone either way.'
              : 'Off: Dataverse stops at the first failed request or change set. On: it runs the rest and returns every result. A change set is all-or-nothing either way.'}
          </Caption1>
        </div>
      </div>

      {em && setReturnResponses && (
        <div
          style={{
            display: 'flex',
            gap: 12,
            alignItems: 'flex-start',
            padding: 12,
            border: `1px solid ${tokens.colorNeutralStroke2}`,
            borderRadius: 6,
            background: tokens.colorNeutralBackground1,
            marginBottom: 16,
          }}
        >
          <Switch
            checked={returnResponses}
            onChange={(_, d) => setReturnResponses(d.checked)}
            aria-label="ReturnResponses"
          />
          <div>
            <strong>ReturnResponses</strong>
            <Caption1 block style={{ color: tokens.colorNeutralForeground3 }}>
              True: a response for every request. False: only the faults, which is what you want for
              bulk loads. A $batch always returns every response; DRS filters the Results the same
              way and sets it in the C#.
            </Caption1>
          </div>
        </div>
      )}

      {needsConsent && (
        <MessageBar intent="info" layout="multiline" style={{ marginBottom: 12 }}>
          <MessageBarBody>
            <MessageBarTitle>Headers need your approval in Power Platform ToolBox</MessageBarTitle>
            This batch sends request headers (per-request headers or continue-on-error). If you
            haven’t approved additional request headers for this tool yet, ToolBox asks first; if
            you reject, nothing is sent.
          </MessageBarBody>
        </MessageBar>
      )}

      {plan.length > 1 && (
        <MessageBar intent="info" layout="multiline" style={{ marginBottom: 12 }}>
          <MessageBarBody>
            <MessageBarTitle>Sent as {plan.length} host calls</MessageBarTitle>
            Power Platform ToolBox sends one change set or one run of independent requests per call,
            so this batch goes out as {plan.length} $batch requests in order (
            {plan.map((s) => (s.transaction ? 'change set' : 'independent')).join(' → ')}). Each
            change set is still atomic.
            {continueOnError ? '' : ' A failed call stops the ones after it.'} The Code tab shows
            the single-envelope form for other clients.
          </MessageBarBody>
        </MessageBar>
      )}

      <div style={{ fontSize: 13, marginBottom: 8 }}>
        <strong>
          {total} request{total === 1 ? '' : 's'}
        </strong>{' '}
        in {groups.length} group{groups.length === 1 ? '' : 's'}
      </div>
      {groups.map((g, gi) => (
        <div
          key={g.id}
          style={{
            border: `1px solid ${tokens.colorNeutralStroke2}`,
            borderLeft: `3px solid ${g.atomic ? 'var(--color-request-batch)' : tokens.colorNeutralStroke1}`,
            borderRadius: 6,
            padding: '8px 12px',
            marginBottom: 8,
            background: tokens.colorNeutralBackground1,
          }}
        >
          <Button
            appearance="transparent"
            onClick={() => onSelect(`group:${g.id}`)}
            style={{ padding: 0 }}
          >
            <strong>
              {em ? 'Requests' : g.atomic ? `Change set ${gi + 1}` : `Independent group ${gi + 1}`}
            </strong>
          </Button>
          {g.items.length === 0 && (
            <Caption1 block style={{ color: tokens.colorNeutralForeground3 }}>
              Empty — skipped when sent.
            </Caption1>
          )}
          {g.items.map((it) => {
            n += 1;
            const sum = itemSummary(it, orgUrl);
            return (
              <div
                key={it.id}
                role="button"
                tabIndex={0}
                onClick={() => onSelect(`item:${it.id}`)}
                onKeyDown={(e) => (e.key === 'Enter' || e.key === ' ') && onSelect(`item:${it.id}`)}
                style={{
                  display: 'flex',
                  gap: 8,
                  alignItems: 'center',
                  padding: '3px 0',
                  cursor: 'pointer',
                }}
              >
                <code style={{ width: 36, color: tokens.colorNeutralForeground3 }}>${n}</code>
                <MethodPill method={sum.method} size="sm" />
                <code
                  style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}
                >
                  {sum.path || '—'}
                </code>
                {it.label && (
                  <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>{it.label}</Caption1>
                )}
              </div>
            );
          })}
        </div>
      ))}
    </div>
  );
}

export function BatchGroupPane({
  group,
  index,
  count,
  contentIdOf,
  orgUrl,
  onSetAtomic,
  onAdd,
  onRemoveGroup,
  onMoveGroup,
  onSelect,
  variant = 'batch',
}: {
  variant?: 'batch' | 'execute-multiple';
  group: BatchGroup;
  index: number;
  count: number;
  contentIdOf: (itemId: string) => string;
  orgUrl: string;
  onSetAtomic: (atomic: boolean) => void;
  onAdd: (op: BatchOpKind) => void;
  onRemoveGroup: () => void;
  onMoveGroup: (delta: -1 | 1) => void;
  onSelect: (id: string) => void;
}) {
  const hasGet = group.items.some((i) => itemSummary(i, orgUrl).method === 'GET');
  const em = variant === 'execute-multiple';
  return (
    <div style={{ maxWidth: 860 }}>
      <PaneHead
        icon={group.atomic ? LockClosed20Regular : TextBulletListSquare20Regular}
        title={
          em
            ? 'Requests'
            : group.atomic
              ? `Change set ${index + 1}`
              : `Independent group ${index + 1}`
        }
        sub={
          em
            ? 'ExecuteMultiple runs each request on its own, in order. It isn’t transactional.'
            : group.atomic
              ? 'Atomic: if any request fails, Dataverse rolls back the whole change set. GET isn’t allowed here.'
              : 'Each request runs on its own; a failure doesn’t undo the others.'
        }
        group="batch"
      >
        {!em && (
          <>
            <Button
              icon={<ArrowUp20Regular />}
              appearance="subtle"
              aria-label="Move group up"
              disabled={index === 0}
              onClick={() => onMoveGroup(-1)}
            />
            <Button
              icon={<ArrowDown20Regular />}
              appearance="subtle"
              aria-label="Move group down"
              disabled={index === count - 1}
              onClick={() => onMoveGroup(1)}
            />
            <Button icon={<Delete20Regular />} appearance="subtle" onClick={onRemoveGroup}>
              Remove group
            </Button>
          </>
        )}
      </PaneHead>

      <div
        style={{
          display: 'flex',
          gap: 12,
          alignItems: 'center',
          marginBottom: 16,
          flexWrap: 'wrap',
        }}
      >
        {!em && (
          <SegmentedToggle ariaLabel="Group kind">
            <ToggleButton checked={group.atomic} onClick={() => onSetAtomic(true)}>
              Change set (atomic)
            </ToggleButton>
            <ToggleButton checked={!group.atomic} onClick={() => onSetAtomic(false)}>
              Independent
            </ToggleButton>
          </SegmentedToggle>
        )}
        <AddRequestMenu onAdd={onAdd} disabledGet={group.atomic} />
      </div>
      {group.atomic && hasGet && (
        <MessageBar intent="error" style={{ marginBottom: 12 }}>
          <MessageBarBody>
            This change set holds a GET request. Move it to an independent group.
          </MessageBarBody>
        </MessageBar>
      )}
      {group.items.length === 0 && (
        <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>No requests yet.</Caption1>
      )}
      {group.items.map((it) => {
        const sum = itemSummary(it, orgUrl);
        return (
          <div
            key={it.id}
            role="button"
            tabIndex={0}
            onClick={() => onSelect(`item:${it.id}`)}
            onKeyDown={(e) => (e.key === 'Enter' || e.key === ' ') && onSelect(`item:${it.id}`)}
            style={{
              display: 'flex',
              gap: 8,
              alignItems: 'center',
              padding: '6px 8px',
              borderBottom: `1px solid ${tokens.colorNeutralStroke3}`,
              cursor: 'pointer',
            }}
          >
            <code style={{ width: 36, color: tokens.colorNeutralForeground3 }}>
              ${contentIdOf(it.id)}
            </code>
            <MethodPill method={sum.method} size="sm" />
            <code style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
              {sum.path || '—'}
            </code>
            {it.label && (
              <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>{it.label}</Caption1>
            )}
          </div>
        );
      })}
    </div>
  );
}
