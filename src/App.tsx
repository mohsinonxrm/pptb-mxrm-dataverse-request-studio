import { useEffect, useState, useSyncExternalStore } from 'react';
import {
  Button,
  FluentProvider,
  MessageBar,
  MessageBarActions,
  MessageBarBody,
  MessageBarTitle,
} from '@fluentui/react-components';
import { Dismiss20Regular } from '@fluentui/react-icons';
import { useStudioStyles } from './primitives/styles';
import { studioLight, studioDark, applyThemeAttr, type ThemeMode } from './theme/theme';
import { FrameHeader } from './shell/FrameHeader';
import { AppDialogHost } from './shell/dialogs';
import { SaveContextRoot, type SaveContextValue } from './state/SaveContext';
import { findRequestType } from './registry/requestTypes';
import { RetrieveMultipleMode } from './modes/RetrieveMultipleMode';
import { RetrieveSingleMode } from './modes/RetrieveSingleMode';
// Less-used modes load on first use (#22).
import {
  RetrieveNextLinkMode,
  PredefinedQueryMode,
  MergeMode,
  AssociateMode,
  DisassociateMode,
  ExecuteWorkflowMode,
  ManageFileMode,
  ManageImageMode,
} from './modes/lazyModes';
import { CreateMode } from './modes/CreateMode';
import { UpdateMode } from './modes/UpdateMode';
import { UpsertMode } from './modes/UpsertMode';
import { DeleteMode } from './modes/DeleteMode';
import { ExecuteActionMode } from './modes/ExecuteActionMode';
import { ExecuteFunctionMode } from './modes/ExecuteFunctionMode';
import { ManageAttachmentMode } from './modes/ManageAttachmentMode';
import { ErrorBoundary } from './shell/ErrorBoundary';
import { clearModeDrafts } from './state/modeDrafts';
import { StubMode } from './modes/StubMode';
import { findTable } from './mock/metadata';
import { _setTableLookup } from './editors/filter/filterTree';
import { HostProvider, useHostSession } from './host/HostContext';
import { SettingsProvider } from './host/usePersistedSettings';
import { OpenSettingsProvider } from './host/useOpenSettings';
import { activeOrg, onActiveOrgChange } from './host/activeOrg';

// Wire metadata lookup into the filter-tree encoder (avoids a circular import).
_setTableLookup(findTable);

// ── Connection switch → fresh modes (#93) ─────────────────────────────
//
// PPTB can move a running tool to another connection, and every later host
// call (queryData, delete, update, …) goes to the new environment. Mode
// state — picked record ids, a ticked Delete confirmation, field values,
// snapshots, results, recent runs — lives inside the mode components, so
// without a reset a Delete or Update prepared against a sandbox would run
// against Prod after the switch (sandbox copies share record GUIDs).
//
// The epoch below bumps whenever the org changes away from a connected org
// (A → B, or A → disconnected). ModeRouter is keyed by it, so the active
// mode remounts with its initial state. The first connection after boot
// ('' → A) doesn't bump: nothing could have run before it, and a remount
// there would only repeat the mode's start-up work.
let orgEpoch = 0;
let lastOrg = activeOrg();
let lastSwitch: { from: string; to: string } | null = null;
onActiveOrgChange((org) => {
  if (lastOrg) {
    orgEpoch++;
    lastSwitch = { from: lastOrg, to: org };
  }
  lastOrg = org;
});
const getOrgEpoch = () => orgEpoch;
function useOrgEpoch(): number {
  return useSyncExternalStore(onActiveOrgChange, getOrgEpoch);
}

export function App() {
  return (
    <HostProvider>
      <SettingsProvider>
        <ThemedApp />
      </SettingsProvider>
    </HostProvider>
  );
}

function ThemedApp() {
  const host = useHostSession();
  // Standalone mode: local toggle in the header controls theme.
  // Embedded mode: host pushes the theme via `settings:updated` events; ignore the local toggle.
  const [localTheme, setLocalTheme] = useState<ThemeMode>('light');
  const themeMode: ThemeMode = host.embedded ? host.theme : localTheme;
  const [activeId, setActiveId] = useState('retrieve-multiple');

  useEffect(() => {
    applyThemeAttr(themeMode);
  }, [themeMode]);

  const fluentTheme = themeMode === 'dark' ? studioDark : studioLight;

  return (
    <FluentProvider theme={fluentTheme} style={{ height: '100vh' }}>
      <OpenSettingsProvider>
        <Frame
          themeMode={themeMode}
          setThemeMode={setLocalTheme}
          activeId={activeId}
          setActiveId={setActiveId}
        />
      </OpenSettingsProvider>
      {/* In-app confirm dialogs and notices; replaces the native confirm/alert boxes (#111). */}
      <AppDialogHost />
    </FluentProvider>
  );
}

function Frame({
  themeMode,
  setThemeMode,
  activeId,
  setActiveId,
}: {
  themeMode: ThemeMode;
  setThemeMode: (m: ThemeMode) => void;
  activeId: string;
  setActiveId: (id: string) => void;
}) {
  const s = useStudioStyles();
  const type = findRequestType(activeId);

  // Single slot for the currently-active mode's save context. The mode
  // publishes via `usePublishSaveContext` and clears on unmount; the
  // top-right Save button in FrameHeader reads via `useSaveContext`.
  // Mounted ABOVE both so the bridge stays alive across mode switches.
  const [saveCtx, setSaveCtx] = useState<SaveContextValue | null>(null);

  // Remount the active mode (keeping which mode is selected) when the
  // connection switches, and say why its inputs were cleared.
  const orgEpoch = useOrgEpoch();
  const [noticeDismissedAt, setNoticeDismissedAt] = useState(orgEpoch);

  return (
    // The active mode's group drives the v9 brand-color cascade defined
    // in index.html. Every primary button,
    // active tab underline, focus ring, and tinted background reads as the
    // group color (green for Write, orange for Execute, etc.). The Read
    // group keeps v9 defaults (blue), so we omit data-mode-group there.
    <div className={s.frame} data-mode-group={type.group === 'read' ? undefined : type.group}>
      <SaveContextRoot value={saveCtx} setValue={setSaveCtx}>
        <FrameHeader
          themeMode={themeMode}
          setThemeMode={setThemeMode}
          activeId={activeId}
          setActiveId={setActiveId}
        />
        {/* Every mode in the registry is now `implemented: true`, so the
            previous `__stub__` prefix detour
            was dead code. The default fallback (line below) still catches
            unknown IDs via StubMode. The `implemented` field is kept on the
            schema as a feature-flag for future modes added in WIP. */}
        <ModeRouter
          key={orgEpoch}
          modeId={activeId}
          themeMode={themeMode}
          setActiveId={setActiveId}
        />
        {orgEpoch !== noticeDismissedAt && lastSwitch && (
          <ConnectionSwitchNotice
            {...lastSwitch}
            modeLabel={type.name}
            onDismiss={() => setNoticeDismissedAt(orgEpoch)}
          />
        )}
      </SaveContextRoot>
    </div>
  );
}

function ConnectionSwitchNotice({
  from,
  to,
  modeLabel,
  onDismiss,
}: {
  from: string;
  to: string;
  modeLabel: string;
  onDismiss: () => void;
}) {
  // Floats over the mode: the frame is a fixed three-row grid that the
  // mode's own body and URL bar fill.
  return (
    <div
      role="status"
      style={{
        position: 'fixed',
        top: 56,
        left: '50%',
        transform: 'translateX(-50%)',
        width: 'min(640px, calc(100vw - 32px))',
        zIndex: 1000,
      }}
    >
      <MessageBar layout="multiline" intent="warning">
        <MessageBarBody>
          <MessageBarTitle>{to ? 'Connection switched' : 'Connection closed'}</MessageBarTitle>
          {to ? (
            <>
              DRS is now connected to <code>{to}</code>. The {modeLabel} request was reset — record
              ids, confirmations, field values and results from <code>{from}</code> were cleared so
              nothing prepared there runs here. Saved requests are kept per connection.
            </>
          ) : (
            <>
              The connection to <code>{from}</code> is gone. The {modeLabel} request was reset so
              nothing prepared there runs against the next connection.
            </>
          )}
        </MessageBarBody>
        <MessageBarActions
          containerAction={
            <Button
              appearance="transparent"
              icon={<Dismiss20Regular />}
              aria-label="Dismiss"
              onClick={onDismiss}
            />
          }
        />
      </MessageBar>
    </div>
  );
}

function ModeRouter({
  modeId,
  themeMode,
  setActiveId,
}: {
  modeId: string;
  themeMode: ThemeMode;
  setActiveId: (id: string) => void;
}) {
  // One boundary around whichever mode is routed (#19): a crash shows a
  // fallback in the mode's place while the header stays usable. It lives
  // inside ModeRouter, so anything that remounts the router via `key` also
  // resets it; switching modes (`resetKey`) clears a caught error.
  const type = findRequestType(modeId);
  return (
    <ErrorBoundary
      scope="mode"
      label={type.id === modeId ? type.name : modeId}
      resetKey={modeId}
      onReload={clearModeDrafts}
    >
      <ModeSwitch modeId={modeId} themeMode={themeMode} setActiveId={setActiveId} />
    </ErrorBoundary>
  );
}

function ModeSwitch({
  modeId,
  themeMode,
  setActiveId,
}: {
  modeId: string;
  themeMode: ThemeMode;
  setActiveId: (id: string) => void;
}) {
  switch (modeId) {
    case 'retrieve-multiple':
      return <RetrieveMultipleMode themeMode={themeMode} />;
    case 'retrieve-single':
      return <RetrieveSingleMode themeMode={themeMode} />;
    case 'retrieve-nextlink':
      return <RetrieveNextLinkMode themeMode={themeMode} />;
    case 'predefined-query':
      return <PredefinedQueryMode themeMode={themeMode} />;
    case 'create':
      return <CreateMode themeMode={themeMode} />;
    case 'update':
      return <UpdateMode themeMode={themeMode} />;
    case 'upsert':
      return <UpsertMode themeMode={themeMode} />;
    case 'delete':
      return <DeleteMode themeMode={themeMode} />;
    case 'merge':
      return <MergeMode themeMode={themeMode} />;
    case 'associate':
      return <AssociateMode themeMode={themeMode} />;
    case 'disassociate':
      return <DisassociateMode themeMode={themeMode} />;
    // The three Execute Action sub-modes share one component at the same tree
    // position. `key` forces a fresh mount per sub-mode so the selected
    // action, params, result and recents don't leak between them (#56).
    case 'exec-action':
      return <ExecuteActionMode key={modeId} themeMode={themeMode} category="oob" />;
    case 'exec-customapi':
      return <ExecuteActionMode key={modeId} themeMode={themeMode} category="custom-api" />;
    case 'exec-customaction':
      return <ExecuteActionMode key={modeId} themeMode={themeMode} category="custom-action" />;
    case 'exec-function':
      return <ExecuteFunctionMode themeMode={themeMode} />;
    case 'exec-workflow':
      return <ExecuteWorkflowMode themeMode={themeMode} />;
    case 'manage-file':
      return <ManageFileMode themeMode={themeMode} />;
    case 'manage-image':
      return <ManageImageMode themeMode={themeMode} />;
    case 'manage-attachment':
      return <ManageAttachmentMode themeMode={themeMode} />;
    default:
      return <StubMode requestId={modeId} onPick={setActiveId} />;
  }
}
