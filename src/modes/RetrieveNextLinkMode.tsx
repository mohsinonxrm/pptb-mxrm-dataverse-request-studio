// RetrieveNextLink — paste an `@odata.nextLink` URL and follow it.
//
// Design notes:
//
//   1. The URL is OPAQUE. Dataverse encodes paging state in a `$skiptoken`
//      cookie. Mutating any other clause (or even URL-decoding + recoding the
//      skiptoken) invalidates the cursor. So this mode has no clause editor —
//      just paste, validate, execute.
//
//   2. NO save/load. Skiptokens are server-side cursors that go stale. Saving
//      one for re-use later is meaningless. This is the one mode that doesn't
//      publish a SaveContext.
//
//   3. CROSS-ORG SAFETY. Pasted URLs come from anywhere — a Teams chat, a
//      log file, an old browser tab. If the URL's host doesn't match the
//      currently-connected Dataverse env, executing it would either:
//        a. 401/403 (token doesn't match audience), or
//        b. silently hit a DIFFERENT tenant (if the user happens to have
//           multi-tenant auth set up).
//      Both outcomes are confusing. We block execute with a clear error and
//      surface the mismatch loudly.
//
//   4. NO MOCK SAMPLE. v2 had a hardcoded `contoso-dev` URL. Per the
//      "metadata-driven, no mock data" directive we start empty; users paste.

import { useMemo, useState } from 'react';
import {
  Settings20Regular, Settings20Filled,
  LineHorizontal320Regular, LineHorizontal320Filled,
  ChevronRight20Regular, ChevronRight20Filled,
  Open20Regular, ClipboardPaste20Regular, Dismiss20Regular,
} from '@fluentui/react-icons';
import {
  Field, Textarea, MessageBar, MessageBarBody, MessageBarTitle, Caption1,
  tokens, Button, Tooltip,
} from '@fluentui/react-components';
import { Sidebar } from '../shell/Sidebar';
import { MainTabs, type MainTab } from '../shell/MainTabs';
import { UrlBar } from '../shell/UrlBar';
import { ModeShell } from '../shell/ModeShell';
import { PaneHead } from '../editors/PaneHead';
import { PreferEditor, emptyPrefer, preferToHeaderString } from '../editors/PreferEditor';
import { HeadersEditor, defaultReadHeaders, headerItemsToObject } from '../editors/HeadersEditor';
import { CodeView } from '../views/CodeView';
import { ResultsView } from '../views/ResultsView';
import { findRequestType } from '../registry/requestTypes';
import { runtime, type ExecResult } from '../engine/runtime';
import { buildRetrieveNextLink } from '../engine/urlBuilder';
import type { RetrieveNextLinkState, RecentRun } from '../state/readState';
import type { ThemeMode } from '../theme/theme';
import { getEnv } from '../mock/environment';
import type { Advisory } from '../primitives/advisories';

const initialState = (): RetrieveNextLinkState => ({
  url: '',
  prefer: { ...emptyPrefer(), maxpagesize: 50 },
  headers: defaultReadHeaders(),
  dirty: new Set(),
});

type ClauseId = 'url' | 'prefer' | 'headers';

// ── URL analysis ──────────────────────────────────────────────────────
// Centralized so the same parse drives the read-only fragments panel,
// the cross-org check, the advisory drawer, and the execute gate.

interface NextLinkAnalysis {
  ok: boolean;
  url: URL | null;
  host: string;
  path: string;
  /** Entity-set segment, e.g. "accounts" from /api/data/v9.2/accounts. */
  entitySet: string | null;
  /** Decoded skiptoken value, or null if not present. */
  skiptoken: string | null;
  /** All query params as a flat list (preserves order). */
  params: { key: string; value: string }[];
  /** Reasons the URL is unsafe to execute, if any. */
  errors: string[];
}

function analyzeNextLink(raw: string, currentHost: string): NextLinkAnalysis {
  const empty: NextLinkAnalysis = {
    ok: false, url: null, host: '', path: '', entitySet: null,
    skiptoken: null, params: [], errors: [],
  };
  if (!raw.trim()) return empty;
  let url: URL;
  try {
    url = new URL(raw);
  } catch {
    return { ...empty, errors: ['Not a valid URL — paste the full https://… string.'] };
  }
  const errors: string[] = [];
  if (url.protocol !== 'https:') errors.push('Must be https:// — Dataverse rejects plain HTTP.');
  // Web-API path shape: /api/data/v<n>.<n>/<entitySet>[?...]
  const apiMatch = url.pathname.match(/^\/api\/data\/v\d+\.\d+\/([^/?]+)/);
  if (!apiMatch) {
    errors.push('Path doesn\'t look like a Dataverse Web API URL (expected /api/data/v<x>.<x>/<entityset>).');
  }
  // Cross-org check — pasted URL must point at the currently-connected host.
  if (currentHost && url.host && url.host.toLowerCase() !== currentHost.toLowerCase()) {
    errors.push(
      `URL host (${url.host}) doesn't match the connected environment (${currentHost}). ` +
      'Switch connections or paste a URL from the current org.',
    );
  }
  const params: { key: string; value: string }[] = [];
  url.searchParams.forEach((v, k) => params.push({ key: k, value: v }));
  const skiptoken = url.searchParams.get('$skiptoken');
  return {
    ok: errors.length === 0,
    url,
    host: url.host,
    path: url.pathname,
    entitySet: apiMatch ? apiMatch[1] : null,
    skiptoken,
    params,
    errors,
  };
}

export function RetrieveNextLinkMode({ themeMode }: { themeMode: ThemeMode }) {
  const type = findRequestType('retrieve-nextlink');
  const env = getEnv();
  const [state, setState] = useState(initialState);
  const [activeNode, setActiveNode] = useState<ClauseId>('url');
  const [tab, setTab] = useState<MainTab>('builder');
  const [result, setResult] = useState<ExecResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [recents, setRecents] = useState<RecentRun[]>([]);

  const built = useMemo(() => buildRetrieveNextLink(state), [state]);
  const analysis = useMemo(() => analyzeNextLink(state.url, env.host), [state.url, env.host]);

  const markDirty = (id: ClauseId) => setState(s => { const d = new Set(s.dirty); d.add(id); return { ...s, dirty: d }; });
  const set = <K extends keyof RetrieveNextLinkState>(k: K, v: RetrieveNextLinkState[K], dirtyId?: ClauseId) => {
    setState(s => ({ ...s, [k]: v }));
    if (dirtyId) markDirty(dirtyId);
  };

  // Aggregate advisories — surfaced via the AdvisoryDrawer in the URL bar.
  // For NextLink, the only signals are URL-shape problems; everything funnels
  // through analyzeNextLink so the drawer chip count == errors list.
  const advisories: Advisory[] = useMemo(() => {
    return analysis.errors.map<Advisory>((msg, i) => ({
      id: `nl-err-${i}`,
      severity: 'error',
      source: 'validation',
      focusNode: 'url',
      title: 'Invalid next-link URL',
      body: msg,
    }));
  }, [analysis]);

  const onExecute = async () => {
    if (!analysis.ok) return;
    setLoading(true);
    const res = await runtime.nextLink(state);
    setResult(res);
    setLoading(false);
    setTab('results');
    setRecents(rs => [{
      id: `r-${Date.now()}`, modeId: 'retrieve-nextlink',
      url: state.url, method: 'GET', ts: Date.now(),
      status: res.status, ms: res.ms,
      rowCount: ((res.body as { value?: unknown[] } | null)?.value?.length) ?? 0,
    }, ...rs].slice(0, 8));
    setState(s => ({ ...s, dirty: new Set() }));
  };

  const disabledReason =
    !state.url.trim() ? 'Paste an @odata.nextLink URL first.' :
    !analysis.ok ? (analysis.errors[0] ?? 'URL is not safe to follow.') :
    null;

  const onPaste = async () => {
    try {
      const text = await navigator.clipboard?.readText();
      if (text) set('url', text.trim(), 'url');
    } catch {
      // Clipboard access denied (e.g. focus required). No-op — user can paste manually.
    }
  };

  const sections = [
    {
      id: 'url', label: 'Target', meta: 'Server-issued',
      items: [{
        id: 'url',
        icon: ChevronRight20Regular, iconFilled: ChevronRight20Filled,
        label: '@odata.nextLink',
        badge: analysis.entitySet ?? (state.url ? 'invalid' : 'empty'),
        badgeAppearance: 'ghost' as const,
        badgeColor: analysis.errors.length ? ('danger' as const) : analysis.entitySet ? ('informative' as const) : ('subtle' as const),
        dirty: state.dirty.has('url'),
      }],
    },
    {
      id: 'prefer', label: 'Prefer',
      items: [{
        id: 'prefer',
        icon: Settings20Regular, iconFilled: Settings20Filled,
        label: 'Prefer header',
        badge: preferToHeaderString(state.prefer) ? 'on' : null, badgeAppearance: 'ghost' as const,
        dirty: state.dirty.has('prefer'),
      }],
    },
    {
      id: 'headers', label: 'Headers', meta: `${state.headers.filter(h => h.enabled).length} active`,
      items: [{
        id: 'headers', icon: LineHorizontal320Regular, iconFilled: LineHorizontal320Filled,
        label: 'HTTP headers',
        badge: state.headers.filter(h => h.enabled).length || null,
        dirty: state.dirty.has('headers'),
      }],
    },
  ];

  let pane: React.ReactNode;
  switch (activeNode) {
    case 'url': pane = (
      <div>
        <PaneHead
          icon={ChevronRight20Filled}
          title="Next-link URL"
          sub="Paste the @odata.nextLink returned by the previous page. The URL is opaque — clauses are not editable here."
        />

        <MessageBar layout="multiline" intent="info" style={{ marginBottom: 12 }}>
          <MessageBarBody>
            <MessageBarTitle>Why no clause editor?</MessageBarTitle>
            The next page must use the <strong>same shape</strong> as the original query. Dataverse
            encodes paging state in a <code>$skiptoken</code> cookie — editing any other clause
            invalidates the cursor.
          </MessageBarBody>
        </MessageBar>

        <div style={{ display: 'flex', gap: 8, marginBottom: 8 }}>
          <Tooltip content="Paste from clipboard" relationship="label">
            <Button
              icon={<ClipboardPaste20Regular />}
              appearance="outline"
              size="small"
              onClick={onPaste}
            >
              Paste
            </Button>
          </Tooltip>
          <Tooltip content="Clear the URL field" relationship="label">
            <Button
              icon={<Dismiss20Regular />}
              appearance="subtle"
              size="small"
              disabled={!state.url}
              onClick={() => set('url', '', 'url')}
            >
              Clear
            </Button>
          </Tooltip>
        </div>

        <Field label="@odata.nextLink" hint="The full URL including https://. Skiptoken cookies are URL-encoded inside.">
          <Textarea
            rows={6}
            value={state.url}
            onChange={(_, d) => set('url', d.value, 'url')}
            placeholder={`https://${env.host}/api/data/v9.2/<entityset>?$select=…&$skiptoken=…`}
            style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 11 }}
          />
        </Field>

        {/* Errors stack — one MessageBar per error so each is independently readable.
            Mirrors what the URL-bar AdvisoryDrawer shows; redundancy is intentional
            because the pane is where users land when they click an advisory's "Open". */}
        {analysis.errors.map((err, i) => (
          <MessageBar key={i} intent="error" layout="multiline" style={{ marginTop: 12 }}>
            <MessageBarBody>
              <MessageBarTitle>Can&apos;t follow this URL</MessageBarTitle>
              {err}
            </MessageBarBody>
          </MessageBar>
        ))}

        {analysis.url && analysis.ok && (
          <div style={{
            marginTop: 16,
            padding: 12,
            border: `1px solid ${tokens.colorNeutralStroke2}`,
            borderRadius: tokens.borderRadiusMedium,
            background: tokens.colorNeutralBackground1,
          }}>
            <Caption1 style={{
              display: 'block',
              fontWeight: 600,
              marginBottom: 8,
              color: tokens.colorNeutralForeground2,
            }}>
              Parsed (read-only)
            </Caption1>
            <div style={{
              display: 'grid',
              gridTemplateColumns: '120px 1fr',
              rowGap: 4,
              columnGap: 12,
              fontFamily: tokens.fontFamilyMonospace,
              fontSize: 11,
            }}>
              <span style={{ color: tokens.colorNeutralForeground3 }}>host</span>
              <span>{analysis.host}</span>
              <span style={{ color: tokens.colorNeutralForeground3 }}>entityset</span>
              <span>{analysis.entitySet ?? '—'}</span>
              {analysis.params.map((p, i) => (
                <span key={i} style={{ display: 'contents' }}>
                  <span style={{ color: tokens.colorBrandForeground1 }}>{p.key}</span>
                  <span style={{ wordBreak: 'break-all' }}>{decodeURIComponent(p.value)}</span>
                </span>
              ))}
            </div>
          </div>
        )}

        {analysis.ok && (
          <div style={{ marginTop: 16 }}>
            <Tooltip content="Open the parsed URL in a new tab (auth will be required)" relationship="description">
              <Button icon={<Open20Regular />} appearance="outline" onClick={() => window.open(state.url, '_blank', 'noopener')}>
                Open in new tab
              </Button>
            </Tooltip>
          </div>
        )}
      </div>
    ); break;
    case 'prefer':  pane = <PreferEditor spec={state.prefer} setSpec={p => set('prefer', p, 'prefer')} />; break;
    case 'headers': pane = <HeadersEditor items={state.headers} setItems={h => set('headers', h, 'headers')} />; break;
  }

  const headersMap = headerItemsToObject(state.headers, preferToHeaderString(state.prefer));
  const codeInputs = { method: 'GET', built, headers: headersMap, isNextLink: true, rawNextLink: state.url };

  return (
    <ModeShell
      sidebar={
        <Sidebar
          type={type}
          urlPreview={built.relativeNoBase}
          sections={sections}
          activeNode={activeNode}
          onSelect={(id) => setActiveNode(id as ClauseId)}
          recents={recents}
        />
      }
      urlBar={
        <UrlBar
          method="GET"
          url={built.relativeUrl}
          executeVerb={type.executeVerb}
          disabledReason={disabledReason}
          loading={loading}
          onExecute={onExecute}
          advisories={advisories}
          onAdvisoryFocus={(nodeId) => setActiveNode(nodeId as ClauseId)}
        />
      }
    >
      <MainTabs tab={tab} onTabChange={setTab} resultCount={(result?.body as { value?: unknown[] } | null)?.value?.length ?? null}>
        {tab === 'builder' && pane}
        {tab === 'code'    && <CodeView themeMode={themeMode} inputs={codeInputs} />}
        {tab === 'results' && <ResultsView result={result} mode="next" />}
      </MainTabs>
    </ModeShell>
  );
}
