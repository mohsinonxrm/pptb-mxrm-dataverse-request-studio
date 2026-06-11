import {
  Button,
  Input,
  Switch,
  tokens,
  Caption1,
  mergeClasses,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
} from '@fluentui/react-components';
import { Add20Regular, Delete20Regular, LineHorizontal320Filled } from '@fluentui/react-icons';
import { useStudioStyles } from '../primitives/styles';
import { PaneHead } from './PaneHead';
import type { RequestGroup } from '../registry/requestTypes';
import { useHostSession } from '../host/HostContext';

export interface HeaderItem {
  id: string;
  name: string;
  value: string;
  enabled: boolean;
  /** built-in, can be enabled/disabled but not removed */
  builtin?: boolean;
  hint?: string;
}

let __hid = 1;
const newHid = () => `h${__hid++}`;

export const defaultReadHeaders = (): HeaderItem[] => [
  {
    id: newHid(),
    name: 'Accept',
    value: 'application/json',
    enabled: true,
    builtin: true,
    hint: 'OData JSON contract — required.',
  },
  { id: newHid(), name: 'OData-MaxVersion', value: '4.0', enabled: true, builtin: true },
  { id: newHid(), name: 'OData-Version', value: '4.0', enabled: true, builtin: true },
  {
    id: newHid(),
    name: 'If-None-Match',
    value: 'null',
    enabled: false,
    hint: 'For Retrieve Single — returns 304 if etag matches.',
  },
  {
    id: newHid(),
    name: 'MSCRMCallerID',
    value: '00000000-0000-0000-0000-000000000000',
    enabled: false,
    hint: 'Impersonation — caller must hold the prvActOnBehalfOfAnotherUser privilege.',
  },
];

/**
 * Default headers for the Write group (Create / Update / Upsert / Delete / Merge).
 *
 * Per https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/create-entity-web-api
 * a Create request requires `Content-Type: application/json; charset=utf-8` in addition
 * to the standard OData trio. `MSCRM.SuppressDuplicateDetection: false` is exposed as
 * an opt-in built-in so the user can enable duplicate detection without having to
 * remember the header name.
 */
export const defaultWriteHeaders = (): HeaderItem[] => [
  {
    id: newHid(),
    name: 'Accept',
    value: 'application/json',
    enabled: true,
    builtin: true,
    hint: 'OData JSON contract — required.',
  },
  {
    id: newHid(),
    name: 'Content-Type',
    value: 'application/json; charset=utf-8',
    enabled: true,
    builtin: true,
    hint: 'POST body content type — required for writes.',
  },
  { id: newHid(), name: 'OData-MaxVersion', value: '4.0', enabled: true, builtin: true },
  { id: newHid(), name: 'OData-Version', value: '4.0', enabled: true, builtin: true },
  {
    id: newHid(),
    name: 'MSCRM.SuppressDuplicateDetection',
    value: 'false',
    enabled: false,
    hint: 'Set to false to enable duplicate detection (default behavior suppresses it). Requires duplicate-detection rules to be configured for the entity.',
  },
  {
    id: newHid(),
    name: 'MSCRMCallerID',
    value: '00000000-0000-0000-0000-000000000000',
    enabled: false,
    hint: 'Impersonation — caller must hold the prvActOnBehalfOfAnotherUser privilege.',
  },
];

export function HeadersEditor({
  items,
  setItems,
  group = 'read',
}: {
  items: HeaderItem[];
  setItems: (items: HeaderItem[]) => void;
  group?: RequestGroup;
}) {
  const s = useStudioStyles();
  const host = useHostSession();
  const update = (id: string, p: Partial<HeaderItem>) =>
    setItems(items.map((it) => (it.id === id ? { ...it, ...p } : it)));
  const remove = (id: string) => setItems(items.filter((it) => it.id !== id));
  const add = () => setItems([...items, { id: newHid(), name: '', value: '', enabled: true }]);

  const activeCount = items.filter((it) => it.enabled && it.name && it.value).length;

  return (
    <div>
      <PaneHead
        icon={LineHorizontal320Filled}
        title="Headers"
        sub={`${activeCount} active. The Prefer header is composed in the Prefer pane.`}
        group={group}
      />

      {/* PPTB doesn't pass our headers through — its Dataverse API client builds
          its own (Authorization, Accept, OData-Version, etc.). The values below
          are still useful as a reference + for standalone testing. */}
      <MessageBar
        layout="multiline"
        intent={host.embedded ? 'warning' : 'info'}
        style={{ marginBottom: 14 }}
      >
        <MessageBarBody>
          <MessageBarTitle>Headers are managed by the PPTB host at execute time.</MessageBarTitle>
          {host.embedded ? (
            <>
              You're running inside Power Platform ToolBox. When you click <strong>Execute</strong>,
              the host's Dataverse API client takes over and supplies its own{' '}
              <code>Authorization</code>, <code>Accept</code>, <code>OData-Version</code>, and other
              defaults — <strong>the values below are not transmitted</strong>. You can still
              configure them here for reference or for the generated code in the Code tab.
            </>
          ) : (
            <>
              When this tool ships into PPTB, the host's Dataverse API client builds its own headers
              (<code>Authorization</code>, <code>Accept</code>, <code>OData-Version</code>, …) and
              the values below will be ignored at execute time. Configure them here anyway — they
              flow through into the generated <strong>Code</strong> tab snippets, which you'd run
              outside PPTB with your own bearer token.
            </>
          )}
        </MessageBarBody>
      </MessageBar>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 6, maxWidth: 760 }}>
        {items.map((it) => (
          <div
            key={it.id}
            className={mergeClasses(s.inlineCard)}
            style={{
              padding: '6px 10px',
              display: 'grid',
              gridTemplateColumns: '40px 220px 1fr 28px',
              gap: 8,
              alignItems: 'center',
            }}
          >
            <Switch
              checked={it.enabled}
              onChange={(_, d) => update(it.id, { enabled: d.checked })}
            />
            <Input
              size="small"
              value={it.name}
              onChange={(_, d) => update(it.id, { name: d.value })}
              placeholder="Header name"
              disabled={it.builtin}
              style={{ fontFamily: tokens.fontFamilyMonospace }}
            />
            <Input
              size="small"
              value={it.value}
              onChange={(_, d) => update(it.id, { value: d.value })}
              placeholder="value"
              style={{ fontFamily: tokens.fontFamilyMonospace }}
            />
            {it.builtin ? (
              <span
                style={{ fontSize: 10, color: tokens.colorNeutralForeground3, textAlign: 'center' }}
              >
                req
              </span>
            ) : (
              <Button
                size="small"
                appearance="subtle"
                icon={<Delete20Regular />}
                onClick={() => remove(it.id)}
                aria-label="Remove"
              />
            )}
            {it.hint && (
              <Caption1 style={{ gridColumn: '2 / span 2', color: tokens.colorNeutralForeground3 }}>
                {it.hint}
              </Caption1>
            )}
          </div>
        ))}
      </div>
      <div style={{ marginTop: 12 }}>
        <Button icon={<Add20Regular />} appearance="outline" size="small" onClick={add}>
          Add header
        </Button>
      </div>
    </div>
  );
}

export const headerItemsToObject = (
  items: HeaderItem[],
  extraPrefer: string | null,
): Record<string, string> => {
  const out: Record<string, string> = {};
  items
    .filter((it) => it.enabled && it.name && it.value)
    .forEach((it) => {
      out[it.name] = it.value;
    });
  if (extraPrefer) out['Prefer'] = extraPrefer;
  return out;
};
