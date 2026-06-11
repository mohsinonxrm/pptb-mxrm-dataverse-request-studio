// Create Record — POST to /<entitySet> with the user's field set as the JSON body.
//
// Sidebar:
//   • Target — pick the entity set
//   • Field set — every populated column, with required-field guard
//   • Prefer — return=representation toggles 201/body vs 204/header
//   • Return $select — only meaningful when return=representation is on
//   • Headers — Content-Type / OData-* / impersonation / duplicate detection
//
// Grounded in: https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/create-entity-web-api
//
// Notes:
//   • Body shape is built by engine/urlBuilder.buildCreateBody (lookups become
//     `<col>@odata.bind`, multi-select choice becomes a comma-separated string).
//   • Required-field validation gates the Execute button — same rule a real
//     Dataverse POST would enforce.
//   • A read-only JSON preview lives next to the field set so the user can see
//     the wire shape forming as they edit.

import { useMemo, useState } from 'react';
import {
  Table20Regular,
  Table20Filled,
  FormNew20Regular,
  FormNew20Filled,
  TextBulletList20Regular,
  TextBulletList20Filled,
  Settings20Regular,
  Settings20Filled,
  LineHorizontal320Regular,
  LineHorizontal320Filled,
} from '@fluentui/react-icons';
import { Sidebar } from '../shell/Sidebar';
import { MainTabs, type MainTab } from '../shell/MainTabs';
import { UrlBar } from '../shell/UrlBar';
import { ModeShell } from '../shell/ModeShell';
import { TargetEditor } from '../editors/TargetEditor';
import { SelectEditor } from '../editors/SelectEditor';
import { FieldSetEditor, seedRequiredFieldValues } from '../editors/FieldSetEditor';
import { PreferEditor, emptyPrefer, preferToHeaderString } from '../editors/PreferEditor';
import {
  HeadersEditor,
  defaultWriteHeaders,
  headerItemsToObject,
  type HeaderItem,
} from '../editors/HeadersEditor';
import { BypassEditor, summarize as summarizeBypass } from '../editors/BypassEditor';
import { applyBypassToHeaders } from '../engine/bypassHeaders';
import { detectBypassAdvisories } from '../engine/bypassAdvisories';
import { disabledReasonFromAdvisories } from '../primitives/advisories';
import { Warning20Filled } from '@fluentui/react-icons';
import { CodeView } from '../views/CodeView';
import { ResultsView } from '../views/ResultsView';
import { findTable, isLookupLike } from '../mock/metadata';
import { findRequestType } from '../registry/requestTypes';
import { buildCreate, buildCreateBody } from '../engine/urlBuilder';
import { runtime, type ExecResult } from '../engine/runtime';
import {
  defaultBypassOptions,
  type CreateState,
  type LookupFieldValue,
  type CreateFieldValue,
} from '../state/writeState';
import type { RecentRun } from '../state/readState';
import type { ThemeMode } from '../theme/theme';
import { useLiveTable } from '../host/useLiveMetadata';
import { useScopedEntities } from '../host/useScopedEntities';
import {
  serializeCreate,
  deserializeCreate,
  hashState,
  type SavedRequest,
  type SerializedCreateState,
} from '../state/savedRequests';
import { usePublishSaveContext } from '../state/SaveContext';

// ──────────────────────────────────────────────────────────────
// Initial state
// ──────────────────────────────────────────────────────────────
// Empty initial state — the user must pick the target table first, which
// then triggers required-field seeding (see Target onTableChange below).
// Previously this seeded "account" with a hardcoded sample name, which
// was active misinformation against a live org that may not even have
// the account entity provisioned.
const initialState = (): CreateState => ({
  table: '',
  fieldValues: {},
  nullFields: [],
  prefer: { ...emptyPrefer(), formattedValues: true, returnRepresentation: true },
  headers: defaultWriteHeaders(),
  returnSelect: [],
  duplicateDetection: false,
  bypass: defaultBypassOptions(),
  dirty: new Set(),
});

type RootClauseId = 'target' | 'fieldset' | 'prefer' | 'returnselect' | 'headers' | 'bypass';

export function CreateMode({ themeMode }: { themeMode: ThemeMode }) {
  const type = findRequestType('create');
  const [state, setState] = useState(initialState);
  // Live-metadata subscription so child editors re-render when columns/relationships land.
  useLiveTable(state.table || null);
  const { entities } = useScopedEntities();
  const [activePath, setActivePath] = useState<string>('target');
  const [tab, setTab] = useState<MainTab>('builder');
  const [result, setResult] = useState<ExecResult | null>(null);
  const [loading, setLoading] = useState(false);
  const [recents, setRecents] = useState<RecentRun[]>([]);

  // ── Save / Load tracking ──
  const [lastSavedId, setLastSavedId] = useState<string | undefined>(undefined);
  const [lastSavedHash, setLastSavedHash] = useState<string | null>(null);

  const built = useMemo(() => buildCreate(state), [state]);
  const body = useMemo(() => buildCreateBody(state), [state]);
  const tbl = findTable(state.table);

  const markDirty = (id: string) =>
    setState((s) => {
      const d = new Set(s.dirty);
      d.add(id);
      return { ...s, dirty: d };
    });
  const set = <K extends keyof CreateState>(k: K, v: CreateState[K], dirtyId?: string) => {
    setState((s) => ({ ...s, [k]: v }));
    if (dirtyId) markDirty(dirtyId);
  };

  // ── Required-field guard ──
  //
  // Empirically validated against `account` (264 attrs). Two require-
  // levels reach this code:
  //
  //   ApplicationRequired — the maker explicitly marked the column as
  //     "Business Required" in the form/column editor. Dataverse expects
  //     you to provide a value; no server-side default.
  //   SystemRequired — platform-managed required. ALMOST ALWAYS server-
  //     defaulted (ownerid → calling user; statecode/statuscode → 0/1;
  //     accountid → fresh GUID). The user doesn't need to provide one.
  //
  // Old behavior gated Execute on BOTH levels. That produced wrong UX:
  // on a fresh `accounts` form, `ownerid` is SystemRequired + valid for
  // create → DRS demanded the user pick an owner before Execute, but
  // Dataverse would have just assigned them automatically. The maker
  // didn't ask for ownership input; the platform did, and the platform
  // already knows the answer.
  //
  // New behavior: gate ONLY on ApplicationRequired. SystemRequired
  // columns that survive the create filter are server-defaulted; the
  // server fills them in if absent. If a power-user wants to override
  // an auto-default (e.g., assign the new account to a different
  // owner), they enable that field manually — it's still visible in
  // the form, just not gating.
  //
  // (Recommended-level columns are still excluded — those are UI hints
  // in model-driven forms, not server-enforced at all.)
  const missingRequired = useMemo(() => {
    if (!tbl) return [];
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
  }, [tbl, state.fieldValues]);

  const disabledReason = !tbl
    ? 'Pick a target table first.'
    : missingRequired.length > 0
      ? `${missingRequired.length} required field${missingRequired.length === 1 ? '' : 's'} unset: ${missingRequired.slice(0, 3).join(', ')}${missingRequired.length > 3 ? '…' : ''}`
      : state.headers.some((h) => h.enabled && !h.name)
        ? 'Fix empty header name.'
        : Object.keys(body).length === 0
          ? 'Body is empty — set at least one field.'
          : null;

  // ── Execute ──
  const onExecute = async () => {
    setLoading(true);
    const res = await runtime.create(state);
    setResult(res);
    setLoading(false);
    setTab('results');
    setRecents((rs) =>
      [
        {
          id: `r-${Date.now()}`,
          modeId: 'create',
          url: built.relativeUrl,
          method: 'POST',
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

  // ── Headers (auto-toggle MSCRM.SuppressDuplicateDetection + bypass composer) ──
  const headers = useMemo<HeaderItem[]>(() => {
    const baseHeaders = state.headers.map((h) => {
      if (h.name === 'MSCRM.SuppressDuplicateDetection') {
        return { ...h, enabled: state.duplicateDetection, value: 'false' };
      }
      return h;
    });
    return applyBypassToHeaders(baseHeaders, state.bypass);
  }, [state.headers, state.duplicateDetection, state.bypass]);

  // ── Advisories — bypass family ──
  const advisories = useMemo(() => detectBypassAdvisories(state.bypass, 'bypass'), [state.bypass]);
  const bypassBlocker = disabledReasonFromAdvisories(advisories);

  // ── Save / Load ──
  const currentSerialized = useMemo(() => serializeCreate(state), [state]);
  const currentHash = useMemo(() => hashState(currentSerialized), [currentSerialized]);
  const isDirty = lastSavedHash === null ? true : currentHash !== lastSavedHash;

  const onSaved = (saved: SavedRequest) => {
    setLastSavedId(saved.id);
    setLastSavedHash(hashState(saved.state));
  };

  const onLoadSaved = (entry: SavedRequest) => {
    if (entry.modeId !== 'create') return;
    const snap = entry.state as SerializedCreateState;
    if (entities.length > 0 && !entities.some((e) => e.logicalName === snap.table)) {
      window.alert(
        `Can't load "${entry.name}": entity \`${snap.table}\` ` +
          `isn't available in this environment. The solution may have been ` +
          `removed or you may be connected to a different org.`,
      );
      return;
    }
    setState(deserializeCreate(snap));
    setLastSavedId(entry.id);
    setLastSavedHash(hashState(snap));
    setResult(null);
    setActivePath('target');
  };

  // Publish — hidden until table picked, no point persisting empty shell.
  usePublishSaveContext(
    useMemo(() => {
      if (!state.table) return null;
      return {
        state: currentSerialized,
        modeId: 'create' as const,
        dirty: isDirty,
        lastSavedId,
        onSaved,
        onLoadSaved,
      };
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [currentSerialized, isDirty, lastSavedId, state.table]),
  );

  // ── Sidebar config ──
  const populatedCount = Object.keys(state.fieldValues).length;
  const returnRep = state.prefer.returnRepresentation;
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
          // Method already lives on the URL bar pill; showing it as a sidebar
          // badge would duplicate the signal in a different color (sidebar
          // badge is brand, pill is method-color).
          dirty: state.dirty.has('target'),
        },
      ],
    },
    {
      // Write modes use "Field set" as the section label, not "Request body".
      id: 'body',
      label: 'Field set',
      meta: `${populatedCount} field${populatedCount === 1 ? '' : 's'}`,
      items: [
        {
          id: 'fieldset',
          icon: FormNew20Regular,
          iconFilled: FormNew20Filled,
          label: 'Field set',
          badge:
            missingRequired.length > 0
              ? `${missingRequired.length} req unset`
              : `${populatedCount} populated`,
          badgeAppearance: 'tint' as const,
          badgeColor: missingRequired.length > 0 ? ('danger' as const) : ('success' as const),
          dirty: state.dirty.has('fieldset'),
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
      meta: `${headers.filter((h) => h.enabled).length} active`,
      items: [
        {
          id: 'headers',
          icon: LineHorizontal320Regular,
          iconFilled: LineHorizontal320Filled,
          label: 'HTTP headers',
          badge: headers.filter((h) => h.enabled).length || null,
          dirty: state.dirty.has('headers'),
        },
        {
          // Compliance audit B-1..B-5 / P-1..P-3 — bypass family pane.
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
            const newTbl = findTable(t);
            const newSeed = newTbl ? seedRequiredFieldValues(newTbl) : {};
            // Switching/clearing the target invalidates every column-bound
            // clause: fieldValues (body) + returnSelect (response shape).
            // Re-seed required fields for the new entity (if any).
            setState((s) => ({
              ...s,
              table: t,
              fieldValues: newSeed,
              returnSelect: [],
              dirty: new Set(['target', 'fieldset']),
            }));
            setResult(null);
          }}
          group="write"
          sub="Pick the entity set the new row will be written to. Switching the target clears the body."
        />
      );
      break;
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
          purpose="create"
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

  const headersMap = headerItemsToObject(headers, preferToHeaderString(state.prefer));
  const codeInputs = {
    method: 'POST',
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
          method="POST"
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
              operation: 'create',
              table: state.table,
              // Create has no recordId pre-execute — WriteResultCard
              // pulls the new GUID from the response's `OData-EntityId`
              // header. recordName: best-effort from the body we sent
              // (the primary-name field is usually populated for a Create
              // form to be valid).
              recordId: null,
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
