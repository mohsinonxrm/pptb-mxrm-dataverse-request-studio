// Upsert Record — PATCH /<entitySet>(<id>|<altKey=val>) that creates if
// missing and updates if present.
//
// Sidebar:
//   • Target          — pick the entity set
//   • Key             — by GUID, or by an alternate (business) key
//   • Field set       — body (changed values; alt-key cols stay out per docs)
//   • Concurrency     — None (default for true upsert) / If-Match: * / If-None-Match: * / etag
//   • Prefer          — return=representation disambiguates 201 (created) vs 200 (updated)
//   • Return $select  — scopes the echoed row when return=representation
//   • Headers
//
// Grounded in:
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/update-delete-entities-using-web-api#upsert-a-table-row
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/use-upsert-insert-update-record

import { useMemo, useState } from 'react';
import {
  Table20Regular,
  Table20Filled,
  Key20Regular,
  Key20Filled,
  FormNew20Regular,
  FormNew20Filled,
  TextBulletList20Regular,
  TextBulletList20Filled,
  Settings20Regular,
  Settings20Filled,
  LineHorizontal320Regular,
  LineHorizontal320Filled,
  ShieldLock20Regular,
  ShieldLock20Filled,
  Warning20Filled,
} from '@fluentui/react-icons';
import { Sidebar } from '../shell/Sidebar';
import { MainTabs, type MainTab } from '../shell/MainTabs';
import { UrlBar } from '../shell/UrlBar';
import { ModeShell } from '../shell/ModeShell';
import { TargetEditor } from '../editors/TargetEditor';
import { SelectEditor } from '../editors/SelectEditor';
import { FieldSetEditor } from '../editors/FieldSetEditor';
import { AlternateKeyEditor } from '../editors/AlternateKeyEditor';
import { PreferEditor, emptyPrefer, preferToHeaderString } from '../editors/PreferEditor';
import { HeadersEditor, defaultWriteHeaders, headerItemsToObject } from '../editors/HeadersEditor';
import { BypassEditor, summarize as summarizeBypass } from '../editors/BypassEditor';
import { applyBypassToHeaders } from '../engine/bypassHeaders';
import { detectBypassAdvisories } from '../engine/bypassAdvisories';
import { disabledReasonFromAdvisories, type Advisory } from '../primitives/advisories';
import { PreconditionEditor, preconditionToHeader } from '../editors/PreconditionEditor';
import { CodeView } from '../views/CodeView';
import { ResultsView } from '../views/ResultsView';
import { findTable, isLookupLike } from '../mock/metadata';
import { findRequestType } from '../registry/requestTypes';
import { buildUpsert, buildUpsertBody } from '../engine/urlBuilder';
import { runtime, type ExecResult } from '../engine/runtime';
import {
  defaultBypassOptions,
  type UpsertState,
  type CreateFieldValue,
  type LookupFieldValue,
} from '../state/writeState';
import type { RecentRun } from '../state/readState';
import type { ThemeMode } from '../theme/theme';
import { useLiveTable } from '../host/useLiveMetadata';
import { useScopedEntities } from '../host/useScopedEntities';
import {
  serializeUpsert,
  deserializeUpsert,
  hashState,
  type SavedRequest,
  type SerializedUpsertState,
} from '../state/savedRequests';
import { usePublishSaveContext } from '../state/SaveContext';

// Empty initial state — user picks the table first, which then re-seeds
// the key shape (GUID by default; switches to first alternate key once
// the entity is known to have any). No mock fieldValues or returnSelect.
const initialState = (): UpsertState => ({
  table: '',
  key: { kind: 'guid', recordId: null },
  fieldValues: {},
  nullFields: [],
  prefer: { ...emptyPrefer(), formattedValues: true, returnRepresentation: true },
  headers: defaultWriteHeaders(),
  returnSelect: [],
  // 'none' = full upsert (create or update — server decides based on
  // whether the row already exists at the addressed key).
  concurrency: { kind: 'none' },
  bypass: defaultBypassOptions(),
  dirty: new Set(),
});

// GUID validator — matches the Update / Delete approach.
const GUID_RE =
  /^[{(]?[0-9a-fA-F]{8}-?[0-9a-fA-F]{4}-?[0-9a-fA-F]{4}-?[0-9a-fA-F]{4}-?[0-9a-fA-F]{12}[)}]?$/;
const isValidGuid = (s: string | null | undefined): boolean => !!s && GUID_RE.test(s.trim());

type RootClauseId =
  | 'target'
  | 'key'
  | 'fieldset'
  | 'precondition'
  | 'prefer'
  | 'returnselect'
  | 'headers'
  | 'bypass';

export function UpsertMode({ themeMode }: { themeMode: ThemeMode }) {
  const type = findRequestType('upsert');
  const [state, setState] = useState(initialState);
  // Live-metadata subscription so child editors re-render when columns/relationships land.
  useLiveTable(state.table || null);
  const { entities } = useScopedEntities();
  const [activePath, setActivePath] = useState<string>('target');
  const [tab, setTab] = useState<MainTab>('builder');
  const [result, setResult] = useState<ExecResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [recents, setRecents] = useState<RecentRun[]>([]);

  // Save / Load tracking (Save button dirty/clean, library checkmark).
  const [lastSavedId, setLastSavedId] = useState<string | undefined>(undefined);
  const [lastSavedHash, setLastSavedHash] = useState<string | null>(null);

  const built = useMemo(() => buildUpsert(state), [state]);
  const body = useMemo(() => buildUpsertBody(state), [state]);
  const tbl = findTable(state.table);

  const markDirty = (id: string) =>
    setState((s) => {
      const d = new Set(s.dirty);
      d.add(id);
      return { ...s, dirty: d };
    });
  const set = <K extends keyof UpsertState>(k: K, v: UpsertState[K], dirtyId?: string) => {
    setState((s) => ({ ...s, [k]: v }));
    if (dirtyId) markDirty(dirtyId);
  };

  const keyOk = (() => {
    const k = state.key;
    if (k.kind === 'guid') return !!k.recordId;
    const def = tbl?.alternateKeys?.find((d) => d.name === k.keyName);
    if (!def) return false;
    return def.columns.every((c) => (k.keyValues[c] ?? '').trim() !== '');
  })();

  const fieldCount = Object.keys(state.fieldValues).length;

  // Required-field guard — conditional on disposition mode.
  //
  //   create-only (If-None-Match: *) — server WILL create. Enforce same
  //     ApplicationRequired rule as Create mode.
  //   update-only (If-Match: *) — server WILL update an existing row.
  //     Required fields are already on the row; don't gate.
  //   none / etag — true upsert. Server decides at runtime whether to
  //     create or update. Required fields might or might not be needed.
  //     Don't hard-gate; leave it to the server's clear error response.
  //
  // ApplicationRequired only (same call as Create mode). SystemRequired
  // columns that survive the create filter are server-defaulted.
  const missingRequired = useMemo(() => {
    if (!tbl) return [] as string[];
    if (state.concurrency.kind !== 'create-only') return [];
    return tbl.columns
      .filter((c) => c.requiredLevel === 'ApplicationRequired' && c.isValidForCreate !== false)
      .filter((c) => {
        const v = state.fieldValues[c.logicalName];
        if (v == null) return true;
        if (typeof v === 'string' && v === '') return true;
        if (Array.isArray(v) && v.length === 0) return true;
        if (isLookupLike(c)) {
          const lk = v as LookupFieldValue;
          return !lk?.id;
        }
        return false;
      })
      .map((c) => c.logicalName);
  }, [tbl, state.concurrency.kind, state.fieldValues]);

  const disabledReason = !tbl
    ? 'Pick a target table first.'
    : !keyOk
      ? state.key.kind === 'guid'
        ? 'Pick a record (or switch to alternate key).'
        : 'Fill in every alternate-key column.'
      : state.key.kind === 'guid' && state.key.recordId && !isValidGuid(state.key.recordId)
        ? 'Record id is not a valid GUID.'
        : missingRequired.length > 0
          ? `${missingRequired.length} required field${missingRequired.length === 1 ? '' : 's'} unset for create-only: ${missingRequired.slice(0, 3).join(', ')}${missingRequired.length > 3 ? '…' : ''}`
          : Object.keys(body).length === 0
            ? 'Body is empty — set at least one field.'
            : state.concurrency.kind === 'etag' && !state.concurrency.etag.trim()
              ? 'Provide an etag value or switch the concurrency mode.'
              : state.headers.some((h) => h.enabled && !h.name)
                ? 'Fix empty header name.'
                : null;

  const onExecute = async () => {
    setLoading(true);
    const res = await runtime.upsert(state);
    setResult(res);
    setLoading(false);
    setTab('results');
    setRecents((rs) =>
      [
        {
          id: `r-${Date.now()}`,
          modeId: 'upsert',
          url: built.relativeUrl,
          method: 'PATCH',
          ts: Date.now(),
          status: res.status,
          ms: res.ms,
          rowCount: res.ok ? 1 : 0,
        },
        ...rs,
      ].slice(0, 8),
    );
    setState((s) => ({ ...s, dirty: new Set() }));
  };

  // ── Compose effective headers ──
  const effectiveHeaders = useMemo(() => {
    let h = [...state.headers];
    const cc = preconditionToHeader(state.concurrency);
    if (cc) {
      h.push({
        id: `__cc__${cc.name}`,
        name: cc.name,
        value: cc.value,
        enabled: true,
        builtin: true,
        hint: 'Auto-composed from the Concurrency pane.',
      });
    }
    h = applyBypassToHeaders(h, state.bypass);
    return h;
  }, [state.headers, state.concurrency, state.bypass]);

  // ── Advisories — bypass + GUID + alt-key host limitation ──
  const advisories = useMemo<Advisory[]>(() => {
    const out: Advisory[] = [...detectBypassAdvisories(state.bypass, 'bypass')];

    if (state.key.kind === 'guid' && state.key.recordId && !isValidGuid(state.key.recordId)) {
      out.push({
        id: 'ups-bad-guid',
        severity: 'error',
        source: 'validation',
        focusNode: 'key',
        title: 'Invalid record id',
        body: `"${state.key.recordId}" isn't a valid GUID. The upsert will fail at the wire.`,
      });
    }

    if (state.key.kind === 'alternate') {
      // Same family as Delete single-property and Update PUT — the host's
      // dataverseAPI doesn't expose a raw-request hook. DRS authors the
      // correct URL; user copies it out for external execution.
      out.push({
        id: 'ups-altkey-unsupported',
        severity: 'error',
        source: 'validation',
        focusNode: 'key',
        title: "Alternate-key Upsert isn't supported by the PPTB host",
        body: (
          <>
            DRS authors the correct request{' '}
            <code>
              PATCH /{tbl?.entitySetName ?? '&lt;set&gt;'}(&lt;keyCol&gt;=&apos;value&apos;)
            </code>
            , but PPTB&apos;s <code>dataverseAPI.update</code> only accepts GUID addressing —
            there&apos;s no raw-request hook for property-keyed URLs. Copy the URL + body from the
            Code tab and run it from Postman / curl / the JS SDK / Power Automate. Or switch the Key
            pane to <strong>By GUID</strong> if you know the record&apos;s primary key.
          </>
        ),
      });
    }

    return out;
  }, [state.bypass, state.key, tbl]);
  const bypassBlocker = disabledReasonFromAdvisories(advisories);

  // ── Save / Load ──
  const currentSerialized = useMemo(() => serializeUpsert(state), [state]);
  const currentHash = useMemo(() => hashState(currentSerialized), [currentSerialized]);
  const isDirty = lastSavedHash === null ? true : currentHash !== lastSavedHash;

  const onSaved = (saved: SavedRequest) => {
    setLastSavedId(saved.id);
    setLastSavedHash(hashState(saved.state));
  };

  const onLoadSaved = (entry: SavedRequest) => {
    if (entry.modeId !== 'upsert') return;
    const snap = entry.state as SerializedUpsertState;
    if (entities.length > 0 && !entities.some((e) => e.logicalName === snap.table)) {
      window.alert(
        `Can't load "${entry.name}": entity \`${snap.table}\` ` +
          `isn't available in this environment. The solution may have been ` +
          `removed or you may be connected to a different org.`,
      );
      return;
    }
    setState(deserializeUpsert(snap));
    setLastSavedId(entry.id);
    setLastSavedHash(hashState(snap));
    setResult(null);
    setActivePath('target');
  };

  // Publish — hidden until a table is picked.
  usePublishSaveContext(
    useMemo(() => {
      if (!state.table) return null;
      return {
        state: currentSerialized,
        modeId: 'upsert' as const,
        dirty: isDirty,
        lastSavedId,
        onSaved,
        onLoadSaved,
      };
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [currentSerialized, isDirty, lastSavedId, state.table]),
  );

  const returnRep = state.prefer.returnRepresentation;
  const concurrencyBadge =
    state.concurrency.kind === 'update-only'
      ? 'If-Match: *'
      : state.concurrency.kind === 'etag'
        ? 'If-Match: etag'
        : state.concurrency.kind === 'create-only'
          ? 'If-None-Match: *'
          : 'no precondition';
  const keyBadge = state.key.kind === 'guid' ? 'GUID' : 'alt key';

  const sections = [
    {
      id: 'target',
      label: 'Target',
      meta: tbl ? `${tbl.displayName} · ${tbl.entitySetName}` : 'Pick a table',
      items: [
        {
          id: 'target',
          icon: Table20Regular,
          iconFilled: Table20Filled,
          label: tbl?.displayName ?? 'Pick a table',
          // Method already lives on the URL bar pill; no sidebar badge needed.
          dirty: state.dirty.has('target'),
        },
      ],
    },
    {
      id: 'key',
      label: 'Key',
      meta: keyBadge,
      items: [
        {
          id: 'key',
          icon: Key20Regular,
          iconFilled: Key20Filled,
          label: state.key.kind === 'guid' ? 'By GUID' : 'By alternate key',
          badge: keyOk ? '✓' : null,
          badgeAppearance: 'ghost' as const,
          dirty: state.dirty.has('key'),
        },
      ],
    },
    {
      // Write modes use "Field set" as the section label.
      id: 'body',
      label: 'Field set',
      meta: `${fieldCount} field${fieldCount === 1 ? '' : 's'}`,
      items: [
        {
          id: 'fieldset',
          icon: FormNew20Regular,
          iconFilled: FormNew20Filled,
          label: 'Field set',
          badge: fieldCount || null,
          badgeAppearance: 'tint' as const,
          badgeColor: fieldCount > 0 ? ('success' as const) : ('subtle' as const),
          dirty: state.dirty.has('fieldset'),
        },
      ],
    },
    // Unified Precondition — previously split into "Mode" (disposition)
    // and "Advanced" (etag input), both editing the same state.concurrency
    // field. PreconditionEditor folds them: radio cards for the four
    // kinds, inline etag input under the etag card.
    {
      id: 'precondition',
      label: 'Precondition',
      meta:
        state.concurrency.kind === 'none'
          ? 'None — server decides'
          : state.concurrency.kind === 'create-only'
            ? 'Require absent record'
            : state.concurrency.kind === 'update-only'
              ? 'Require existing record'
              : 'Match specific ETag',
      items: [
        {
          id: 'precondition',
          icon: ShieldLock20Regular,
          iconFilled: ShieldLock20Filled,
          label: 'Optimistic concurrency',
          badge:
            state.concurrency.kind === 'none'
              ? 'none'
              : state.concurrency.kind === 'create-only'
                ? 'If-None-Match: *'
                : state.concurrency.kind === 'update-only'
                  ? 'If-Match: *'
                  : 'If-Match: etag',
          badgeAppearance: 'ghost' as const,
          dirty: state.dirty.has('precondition'),
        },
      ],
    },
    {
      id: 'prefer',
      label: 'Prefer',
      meta: returnRep ? 'return=representation' : '204 No Content',
      items: [
        {
          id: 'prefer',
          icon: Settings20Regular,
          iconFilled: Settings20Filled,
          label: 'Prefer header',
          badge: preferToHeaderString(state.prefer) ? 'on' : null,
          badgeAppearance: 'ghost' as const,
          dirty: state.dirty.has('prefer'),
        },
        ...(returnRep
          ? [
              {
                id: 'returnselect',
                icon: TextBulletList20Regular,
                iconFilled: TextBulletList20Filled,
                label: '$select (response)',
                code: true,
                badge: state.returnSelect.length || null,
                dirty: state.dirty.has('returnselect'),
              },
            ]
          : []),
      ],
    },
    {
      id: 'headers',
      label: 'Headers',
      meta: `${state.headers.filter((h) => h.enabled).length} active`,
      items: [
        {
          id: 'headers',
          icon: LineHorizontal320Regular,
          iconFilled: LineHorizontal320Filled,
          label: 'HTTP headers',
          badge: state.headers.filter((h) => h.enabled).length || null,
          dirty: state.dirty.has('headers'),
        },
        {
          id: 'bypass',
          icon: Warning20Filled,
          iconFilled: Warning20Filled,
          label: 'Bypass logic',
          badge: summarizeBypass(state.bypass),
          badgeAppearance:
            state.bypass.businessLogic !== 'none' || state.bypass.suppressFlows
              ? ('tint' as const)
              : ('ghost' as const),
          badgeColor:
            state.bypass.businessLogic !== 'none' || state.bypass.suppressFlows
              ? ('warning' as const)
              : ('subtle' as const),
          dirty: state.dirty.has('bypass'),
        },
      ],
    },
  ];

  // ── Builder pane router ──
  let pane: React.ReactNode;
  const root = activePath as RootClauseId;
  switch (root) {
    case 'target':
      pane = (
        <TargetEditor
          table={state.table}
          onTableChange={(t) => {
            // Switching/clearing the target invalidates the key (alt-key
            // columns belong to the old entity), the body, and the
            // response $select. Re-seed the key for the new entity's
            // first alternate key (or empty GUID mode).
            const newTbl = findTable(t);
            const firstAltKey = newTbl?.alternateKeys?.[0];
            setState((s) => ({
              ...s,
              table: t,
              key: firstAltKey
                ? { kind: 'alternate', keyName: firstAltKey.name, keyValues: {} }
                : { kind: 'guid', recordId: null },
              fieldValues: {},
              returnSelect: [],
              dirty: new Set(['target', 'key', 'fieldset']),
            }));
            setResult(null);
          }}
          group="write"
          sub="Pick the entity set for the upsert. Switching the target clears the key and field set."
        />
      );
      break;
    case 'key':
      pane = (
        <AlternateKeyEditor
          table={state.table}
          keyMode={state.key}
          setKeyMode={(k) => set('key', k, 'key')}
          group="write"
        />
      );
      break;
    // Mode + Advanced collapsed into Precondition — see case 'precondition' below.
    case 'fieldset':
      pane = (
        <FieldSetEditor
          table={state.table}
          values={state.fieldValues}
          setValues={(next) =>
            set('fieldValues', next as Record<string, CreateFieldValue>, 'fieldset')
          }
          nullFields={state.nullFields}
          setNullFields={(n) => set('nullFields', n, 'fieldset')}
          group="write"
          themeMode={themeMode}
          // Intersection of IsValidForCreate AND IsValidForUpdate — the
          // body has to be acceptable on both paths since Upsert decides
          // at runtime which it is (unless concurrency=create-only /
          // update-only, but the picker filter stays conservative).
          purpose="upsert"
        />
      );
      break;
    case 'precondition':
      pane = (
        <PreconditionEditor
          mode={state.concurrency}
          setMode={(m) => set('concurrency', m, 'precondition')}
          // Upsert is the only mode that surfaces all four kinds — the
          // create-only / update-only options control the disposition
          // semantically (server creates vs updates only); 'none' is true
          // upsert (server decides).
          available={['none', 'create-only', 'update-only', 'etag']}
          group="write"
        />
      );
      break;
    case 'prefer':
      pane = (
        <PreferEditor
          spec={state.prefer}
          setSpec={(p) => set('prefer', p, 'prefer')}
          group="write"
        />
      );
      break;
    case 'returnselect':
      pane = (
        <SelectEditor
          table={state.table}
          selectedIds={state.returnSelect}
          setSelectedIds={(ids) => set('returnSelect', ids, 'returnselect')}
          group="write"
        />
      );
      break;
    case 'headers':
      pane = (
        <HeadersEditor
          items={state.headers}
          setItems={(h) => set('headers', h, 'headers')}
          group="write"
        />
      );
      break;
    case 'bypass':
      pane = (
        <BypassEditor
          value={state.bypass}
          onChange={(b) => set('bypass', b, 'bypass')}
          group="write"
        />
      );
      break;
  }

  const headersMap = headerItemsToObject(effectiveHeaders, preferToHeaderString(state.prefer));
  const codeInputs = {
    method: 'PATCH',
    built,
    headers: headersMap,
    body,
    entityLogical: built.entityLogical,
  };

  return (
    <ModeShell
      sidebar={
        <Sidebar
          type={type}
          urlPreview={built.relativeNoBase}
          sections={sections}
          activeNode={activePath}
          onSelect={(id) => {
            setActivePath(id);
            setTab('builder');
          }}
          recents={recents}
        />
      }
      urlBar={
        <UrlBar
          method="PATCH"
          url={built.relativeUrl}
          executeVerb={type.executeVerb}
          disabledReason={disabledReason ?? bypassBlocker}
          loading={loading}
          onExecute={onExecute}
          advisories={advisories}
          onAdvisoryFocus={setActivePath}
        />
      }
    >
      <MainTabs tab={tab} onTabChange={setTab} resultCount={result?.ok ? 1 : null}>
        {tab === 'builder' && pane}
        {tab === 'code' && <CodeView themeMode={themeMode} inputs={codeInputs} />}
        {tab === 'results' && (
          <ResultsView
            result={result}
            mode="single"
            table={state.table}
            writeContext={{
              operation: 'upsert',
              table: state.table,
              // GUID mode → known id. Alt-key mode → unknown until the
              // response comes back (and even then PPTB blocks alt-key
              // via the advisory, so this path rarely fires in practice).
              recordId: state.key.kind === 'guid' ? state.key.recordId : null,
              recordName: (() => {
                if (!tbl) return null;
                const v = state.fieldValues[tbl.primaryName];
                return typeof v === 'string' && v ? v : null;
              })(),
            }}
          />
        )}
      </MainTabs>
    </ModeShell>
  );
}
