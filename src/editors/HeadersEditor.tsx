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
import { hostSupportsAdditionalHeaders } from '../host/pptbBridge';
import { isCredentialHeader } from '../engine/sensitiveHeaders';
import { SENT_HEADERS, isHostOwnedHeader, type HeaderPolicy } from '../engine/requestHeaderPolicy';

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
    value: '',
    enabled: false,
    hint: "Retrieve Single: put the row's @odata.etag here to get 304 Not Modified when it hasn't changed.",
  },
  ...impersonationPresets(),
];

/**
 * Impersonation presets, off by default: by the user's systemuserid, or by
 * their Microsoft Entra object id (#109). Either way the signed-in user needs
 * the prvActOnBehalfOfAnotherUser privilege.
 */
const impersonationPresets = (): HeaderItem[] => [
  {
    id: newHid(),
    name: 'MSCRMCallerID',
    value: '00000000-0000-0000-0000-000000000000',
    enabled: false,
    hint: 'Impersonation by systemuserid — caller must hold the prvActOnBehalfOfAnotherUser privilege.',
  },
  {
    id: newHid(),
    name: 'CallerObjectId',
    value: '00000000-0000-0000-0000-000000000000',
    enabled: false,
    hint: "Impersonation by the user's Microsoft Entra object id — same privilege.",
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
  ...impersonationPresets(),
];

export function HeadersEditor({
  items,
  setItems,
  group = 'read',
  sendsHeaders,
  policy,
}: {
  items: HeaderItem[];
  setItems: (items: HeaderItem[]) => void;
  group?: RequestGroup;
  /** Execute sends these headers on PPTB 1.2.7+. Defaults to the write group;
   *  Bulk Write (#257), Execute Action and Execute Workflow (#266) send them too. */
  sendsHeaders?: boolean;
  /** Execute sends only this policy's allow-list (reads, relate, Execute Function; #109). */
  policy?: HeaderPolicy;
}) {
  const s = useStudioStyles();
  const host = useHostSession();
  const update = (id: string, p: Partial<HeaderItem>) =>
    setItems(items.map((it) => (it.id === id ? { ...it, ...p } : it)));
  const remove = (id: string) => setItems(items.filter((it) => it.id !== id));
  const add = () => setItems([...items, { id: newHid(), name: '', value: '', enabled: true }]);

  const activeCount = items.filter((it) => it.enabled && it.name && it.value).length;
  const canSend = host.embedded && hostSupportsAdditionalHeaders();
  const sent = policy ? SENT_HEADERS[policy] : [];
  /** An enabled row Execute doesn't send under the allow-list (Code tab only). */
  const codeTabOnly = (it: HeaderItem) =>
    !!policy &&
    canSend &&
    it.enabled &&
    !!it.name.trim() &&
    !isHostOwnedHeader(it.name) &&
    !isCredentialHeader(it.name) &&
    !sent.some((n) => n.toLowerCase() === it.name.trim().toLowerCase());

  return (
    <div>
      <PaneHead
        icon={LineHorizontal320Filled}
        title="Headers"
        sub={`${activeCount} active. The Prefer header is composed in the Prefer pane.`}
        group={group}
      />

      {/* PPTB 1.2.7+ sends the write group's headers as `additionalHeaders`
          (minus the ones its Dataverse client builds itself), and so do the
          modes that pass `sendsHeaders` (Bulk Write, Execute Action and
          Execute Workflow, #266). Every other mode still runs with the host's
          own headers. The values below are always emitted into the Code tab. */}
      {policy && canSend ? (
        <MessageBar layout="multiline" intent="info" style={{ marginBottom: 14 }}>
          <MessageBarBody>
            <MessageBarTitle>
              Execute sends {sent.join(', ')} through Power Platform ToolBox.
            </MessageBarTitle>
            {policy === 'read' ? (
              <>
                And a <code>Prefer</code> header when you set a page size, which keeps
                ToolBox&apos;s <code>odata.include-annotations=&quot;*&quot;</code>.{' '}
              </>
            ) : null}
            Other rows below go into the Code tab only. ToolBox builds <code>Authorization</code>,{' '}
            <code>Accept</code> and the <code>OData-*</code> headers itself. If you haven&apos;t
            approved additional request headers for this tool yet, ToolBox asks first; if you
            reject, the request isn&apos;t sent.
          </MessageBarBody>
        </MessageBar>
      ) : (sendsHeaders ?? group === 'write') && canSend ? (
        <MessageBar layout="multiline" intent="info" style={{ marginBottom: 14 }}>
          <MessageBarBody>
            <MessageBarTitle>
              Enabled headers are sent through Power Platform ToolBox.
            </MessageBarTitle>
            When you click <strong>Execute</strong>, the host sends the enabled headers below, plus
            any that this mode's Precondition and Bypass panes add. It builds{' '}
            <code>Authorization</code>, <code>Accept</code>, <code>Content-Type</code>,{' '}
            <code>OData-Version</code>, <code>OData-MaxVersion</code> and <code>Prefer</code>{' '}
            itself, so those rows aren&apos;t sent. If you haven&apos;t approved additional request
            headers for this tool yet, ToolBox asks first; if you reject, the request isn&apos;t
            sent.
          </MessageBarBody>
        </MessageBar>
      ) : (
        <MessageBar
          layout="multiline"
          intent={host.embedded ? 'warning' : 'info'}
          style={{ marginBottom: 14 }}
        >
          <MessageBarBody>
            <MessageBarTitle>Headers are managed by the PPTB host at execute time.</MessageBarTitle>
            {host.embedded ? (
              <>
                You're running inside Power Platform ToolBox. When you click{' '}
                <strong>Execute</strong>, the host's Dataverse API client takes over and supplies
                its own <code>Authorization</code>, <code>Accept</code>, <code>OData-Version</code>,
                and other defaults — <strong>the values below are not transmitted</strong>. You can
                still configure them here for reference or for the generated code in the Code tab.
              </>
            ) : (
              <>
                When this tool ships into PPTB, the host's Dataverse API client builds its own
                headers (<code>Authorization</code>, <code>Accept</code>, <code>OData-Version</code>
                , …) and the values below will be ignored at execute time. Configure them here
                anyway — they flow through into the generated <strong>Code</strong> tab snippets,
                which you'd run outside PPTB with your own bearer token.
              </>
            )}
          </MessageBarBody>
        </MessageBar>
      )}
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
              aria-label={`Send ${it.name || 'header'}`}
              checked={it.enabled}
              onChange={(_, d) => update(it.id, { enabled: d.checked })}
            />
            <Input
              size="small"
              value={it.name}
              onChange={(_, d) => update(it.id, { name: d.value })}
              placeholder="Header name"
              aria-label="Header name"
              disabled={it.builtin}
              style={{ fontFamily: tokens.fontFamilyMonospace }}
            />
            <Input
              size="small"
              value={it.value}
              onChange={(_, d) => update(it.id, { value: d.value })}
              placeholder="value"
              aria-label={`${it.name || 'Header'} value`}
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
            {codeTabOnly(it) && (
              <Caption1 style={{ gridColumn: '2 / span 2', color: tokens.colorNeutralForeground3 }}>
                Code tab only: Execute doesn&apos;t send this header in this mode.
              </Caption1>
            )}
            {isCredentialHeader(it.name) && (
              <Caption1
                style={{
                  gridColumn: '2 / span 2',
                  color: tokens.colorPaletteDarkOrangeForeground1,
                }}
              >
                Not sent, saved or exported. Power Platform ToolBox authenticates every request, and
                the Code tab adds its own token placeholder.
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
