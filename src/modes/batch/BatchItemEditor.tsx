// Compact editor for one $batch request (#9): pick an operation, fill the
// few fields it needs, see the part it produces.

import { useMemo } from 'react';
import {
  Button,
  Caption1,
  Dropdown,
  Field,
  Input,
  MessageBar,
  MessageBarBody,
  Option,
  Switch,
  Textarea,
  Tooltip,
  tokens,
} from '@fluentui/react-components';
import {
  Add20Regular,
  ArrowDown20Regular,
  ArrowUp20Regular,
  Delete20Regular,
  Document20Regular,
} from '@fluentui/react-icons';
import { PaneHead } from '../../editors/PaneHead';
import { useScopedEntities } from '../../host/useScopedEntities';
import { useLiveTable } from '../../host/useLiveMetadata';
import { isCollectionValuedNav } from '../../state/relateState';
import {
  type BatchGroup,
  type BatchItem,
  type BatchIssue,
  type BatchMethod,
  type BatchOpKind,
  type WireRequest,
} from '../../engine/batch/batchModel';
import { encodePart } from '../../engine/batch/multipart';

const OPS: { id: BatchOpKind; label: string }[] = [
  { id: 'create', label: 'Create (POST)' },
  { id: 'update', label: 'Update (PATCH)' },
  { id: 'delete', label: 'Delete (DELETE)' },
  { id: 'retrieve', label: 'Retrieve (GET)' },
  { id: 'associate', label: 'Associate (POST $ref)' },
  { id: 'custom', label: 'Custom request' },
];
const METHODS: BatchMethod[] = ['GET', 'POST', 'PATCH', 'PUT', 'DELETE'];

const HEADER_PRESETS = [
  { name: 'Prefer', value: 'return=representation' },
  { name: 'If-Match', value: '*' },
  { name: 'MSCRM.SuppressDuplicateDetection', value: 'false' },
];

const mono = { fontFamily: tokens.fontFamilyMonospace, fontSize: 12 };

export function BatchItemEditor({
  item,
  contentId,
  group,
  groups,
  wire,
  issues,
  onChange,
  onRemove,
  onMove,
  onMoveToGroup,
}: {
  item: BatchItem;
  contentId: string;
  group: BatchGroup;
  groups: BatchGroup[];
  wire: WireRequest | null;
  issues: BatchIssue[];
  onChange: (patch: Partial<BatchItem>) => void;
  onRemove: () => void;
  onMove: (delta: -1 | 1) => void;
  onMoveToGroup: (groupId: string) => void;
}) {
  const { entities } = useScopedEntities();
  // Resolve the entity set back to a logical name for navigation-property hints.
  const logical = useMemo(
    () => entities.find((e) => e.entitySetName === item.entitySet.trim())?.logicalName ?? null,
    [entities, item.entitySet],
  );
  const { table } = useLiveTable(item.op === 'associate' ? logical : null);
  const navs = useMemo(
    () => (table?.navigationProperties ?? []).filter(isCollectionValuedNav),
    [table],
  );
  const listId = `drs-batch-sets-${item.id}`;
  const navListId = `drs-batch-navs-${item.id}`;
  const usesSet = item.op !== 'custom';
  const usesId = item.op === 'update' || item.op === 'delete' || item.op === 'associate';
  const usesBody = item.op === 'create' || item.op === 'update' || item.op === 'custom';
  const groupIndex = groups.indexOf(group);

  const setHeader = (i: number, patch: Partial<BatchItem['headers'][number]>) =>
    onChange({ headers: item.headers.map((h, j) => (j === i ? { ...h, ...patch } : h)) });

  return (
    <div style={{ maxWidth: 860 }}>
      <PaneHead
        icon={Document20Regular}
        title={`Request $${contentId}`}
        sub={
          group.atomic
            ? `In change set ${groupIndex + 1} — rolls back with the rest of the change set if any request fails.`
            : `In group ${groupIndex + 1} — independent; runs on its own.`
        }
        group="batch"
      >
        <Tooltip content="Move up" relationship="label">
          <Button icon={<ArrowUp20Regular />} appearance="subtle" onClick={() => onMove(-1)} />
        </Tooltip>
        <Tooltip content="Move down" relationship="label">
          <Button icon={<ArrowDown20Regular />} appearance="subtle" onClick={() => onMove(1)} />
        </Tooltip>
        <Button icon={<Delete20Regular />} appearance="subtle" onClick={onRemove}>
          Remove
        </Button>
      </PaneHead>

      {issues.length > 0 && (
        <MessageBar
          intent={issues.some((i) => i.severity === 'error') ? 'error' : 'warning'}
          layout="multiline"
          style={{ marginBottom: 12 }}
        >
          <MessageBarBody>
            {issues.map((i, n) => (
              <div key={n}>{i.message}</div>
            ))}
          </MessageBarBody>
        </MessageBar>
      )}

      <div
        style={{
          display: 'grid',
          gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))',
          gap: 12,
        }}
      >
        <Field label="Operation">
          <Dropdown
            value={OPS.find((o) => o.id === item.op)?.label ?? ''}
            selectedOptions={[item.op]}
            onOptionSelect={(_, d) => {
              const op = d.optionValue as BatchOpKind;
              const patch: Partial<BatchItem> = { op };
              if ((op === 'create' || op === 'update') && !item.body.trim())
                patch.body = '{\n  \n}';
              onChange(patch);
            }}
          >
            {OPS.map((o) => (
              <Option key={o.id} value={o.id}>
                {o.label}
              </Option>
            ))}
          </Dropdown>
        </Field>
        {groups.length > 1 && (
          <Field label="Group">
            <Dropdown
              value={`${group.atomic ? 'Change set' : 'Independent'} ${groupIndex + 1}`}
              selectedOptions={[group.id]}
              onOptionSelect={(_, d) => d.optionValue && onMoveToGroup(d.optionValue)}
            >
              {groups.map((g, i) => (
                <Option key={g.id} value={g.id}>
                  {`${g.atomic ? 'Change set' : 'Independent'} ${i + 1}`}
                </Option>
              ))}
            </Dropdown>
          </Field>
        )}
        <Field label="Label (optional)">
          <Input value={item.label} onChange={(_, d) => onChange({ label: d.value })} />
        </Field>
      </div>

      <div
        style={{
          display: 'grid',
          gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))',
          gap: 12,
          marginTop: 12,
        }}
      >
        {item.op === 'custom' && (
          <>
            <Field label="Method">
              <Dropdown
                value={item.method}
                selectedOptions={[item.method]}
                onOptionSelect={(_, d) => onChange({ method: d.optionValue as BatchMethod })}
              >
                {METHODS.map((m) => (
                  <Option key={m} value={m}>
                    {m}
                  </Option>
                ))}
              </Dropdown>
            </Field>
            <Field
              label="URL (relative to /api/data/v9.2/)"
              hint="e.g. WhoAmI, accounts?$top=5, $1/primarycontactid/$ref"
              style={{ gridColumn: 'span 2' }}
            >
              <Input
                value={item.url}
                onChange={(_, d) => onChange({ url: d.value })}
                input={{ style: mono }}
              />
            </Field>
          </>
        )}
        {usesSet && (
          <Field
            label={item.op === 'associate' ? 'Source table (entity set)' : 'Table (entity set)'}
          >
            <Input
              value={item.entitySet}
              onChange={(_, d) => onChange({ entitySet: d.value })}
              placeholder="accounts"
              input={{ list: listId, style: mono }}
            />
          </Field>
        )}
        {(usesId || item.op === 'retrieve') && (
          <Field
            label={item.op === 'retrieve' ? 'Record id (optional)' : 'Record id'}
            hint={
              item.op === 'retrieve'
                ? undefined
                : 'GUID, alternate key, or $n for a record created earlier in this change set'
            }
          >
            <Input
              value={item.recordId}
              onChange={(_, d) => onChange({ recordId: d.value })}
              placeholder="00000000-0000-0000-0000-000000000000"
              input={{ style: mono }}
            />
          </Field>
        )}
        {item.op === 'retrieve' && (
          <Field label="Query options" hint="Without the leading ?">
            <Input
              value={item.query}
              onChange={(_, d) => onChange({ query: d.value })}
              placeholder="$select=name&$top=10"
              input={{ style: mono }}
            />
          </Field>
        )}
        {item.op === 'associate' && (
          <>
            <Field label="Navigation property (collection-valued)">
              <Input
                value={item.navProperty}
                onChange={(_, d) => onChange({ navProperty: d.value })}
                placeholder="contact_customer_accounts"
                input={{ list: navListId, style: mono }}
              />
            </Field>
            <Field label="Target table (entity set)">
              <Input
                value={item.targetSet}
                onChange={(_, d) => onChange({ targetSet: d.value })}
                placeholder="contacts"
                input={{ list: listId, style: mono }}
              />
            </Field>
            <Field label="Target record id" hint="GUID or $n">
              <Input
                value={item.targetId}
                onChange={(_, d) => onChange({ targetId: d.value })}
                input={{ style: mono }}
              />
            </Field>
          </>
        )}
      </div>
      <datalist id={listId}>
        {entities.map((e) => (
          <option key={e.logicalName} value={e.entitySetName}>
            {e.displayName}
          </option>
        ))}
      </datalist>
      <datalist id={navListId}>
        {navs.map((n) => (
          <option key={n.name} value={n.name}>
            {n.targetEntity}
          </option>
        ))}
      </datalist>

      {usesBody && (
        <Field
          label="Body (JSON)"
          hint={
            group.atomic
              ? 'Reference a record created earlier in this change set with "$n", e.g. "primarycontactid@odata.bind": "$1".'
              : undefined
          }
          style={{ marginTop: 12 }}
        >
          <Textarea
            value={item.body}
            onChange={(_, d) => onChange({ body: d.value })}
            resize="vertical"
            rows={8}
            textarea={{ style: mono, spellCheck: false }}
          />
        </Field>
      )}

      <div style={{ marginTop: 16 }}>
        <div
          style={{
            display: 'flex',
            alignItems: 'center',
            gap: 8,
            marginBottom: 6,
            flexWrap: 'wrap',
          }}
        >
          <strong style={{ fontSize: 13 }}>Request headers</strong>
          <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
            Sent inside this part. Power Platform ToolBox asks you to approve them the first time.
          </Caption1>
        </div>
        {item.headers.map((h, i) => (
          <div key={i} style={{ display: 'flex', gap: 8, marginBottom: 6, alignItems: 'center' }}>
            <Switch
              checked={h.enabled}
              onChange={(_, d) => setHeader(i, { enabled: d.checked })}
              aria-label={`Send ${h.name || 'header'}`}
            />
            <Input
              value={h.name}
              onChange={(_, d) => setHeader(i, { name: d.value })}
              placeholder="Name"
              aria-label="Header name"
              input={{ style: mono }}
              style={{ flex: '1 1 0', minWidth: 0 }}
            />
            <Input
              value={h.value}
              onChange={(_, d) => setHeader(i, { value: d.value })}
              placeholder="Value"
              aria-label="Header value"
              input={{ style: mono }}
              style={{ flex: '2 1 0', minWidth: 0 }}
            />
            <Button
              icon={<Delete20Regular />}
              appearance="subtle"
              aria-label="Remove header"
              onClick={() => onChange({ headers: item.headers.filter((_, j) => j !== i) })}
            />
          </div>
        ))}
        <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
          <Button
            size="small"
            icon={<Add20Regular />}
            onClick={() =>
              onChange({ headers: [...item.headers, { name: '', value: '', enabled: true }] })
            }
          >
            Header
          </Button>
          {HEADER_PRESETS.filter(
            (p) => !item.headers.some((h) => h.name.toLowerCase() === p.name.toLowerCase()),
          ).map((p) => (
            <Button
              key={p.name}
              size="small"
              appearance="subtle"
              onClick={() => onChange({ headers: [...item.headers, { ...p, enabled: true }] })}
            >
              + {p.name}: {p.value}
            </Button>
          ))}
        </div>
      </div>

      {wire && (
        <div style={{ marginTop: 16 }}>
          <strong style={{ fontSize: 13 }}>Part preview</strong>
          <pre
            style={{
              ...mono,
              background: tokens.colorNeutralBackground1,
              border: `1px solid ${tokens.colorNeutralStroke2}`,
              borderRadius: 4,
              padding: 10,
              marginTop: 6,
              whiteSpace: 'pre-wrap',
              wordBreak: 'break-all',
            }}
          >
            {encodePart({ contentId, wire }).replace(/\r\n/g, '\n')}
          </pre>
        </div>
      )}
      {item.source && (
        <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>From: {item.source}</Caption1>
      )}
    </div>
  );
}
