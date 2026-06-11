import { useEffect, useState } from 'react';
import { FluentProvider } from '@fluentui/react-components';
import { useStudioStyles } from './primitives/styles';
import { studioLight, studioDark, applyThemeAttr, type ThemeMode } from './theme/theme';
import { FrameHeader } from './shell/FrameHeader';
import { SaveContextRoot, type SaveContextValue } from './state/SaveContext';
import { findRequestType } from './registry/requestTypes';
import { RetrieveMultipleMode } from './modes/RetrieveMultipleMode';
import { RetrieveSingleMode } from './modes/RetrieveSingleMode';
import { RetrieveNextLinkMode } from './modes/RetrieveNextLinkMode';
import { PredefinedQueryMode } from './modes/PredefinedQueryMode';
import { CreateMode } from './modes/CreateMode';
import { UpdateMode } from './modes/UpdateMode';
import { UpsertMode } from './modes/UpsertMode';
import { DeleteMode } from './modes/DeleteMode';
import { MergeMode } from './modes/MergeMode';
import { AssociateMode } from './modes/AssociateMode';
import { DisassociateMode } from './modes/DisassociateMode';
import { ExecuteActionMode } from './modes/ExecuteActionMode';
import { ExecuteFunctionMode } from './modes/ExecuteFunctionMode';
import { ExecuteWorkflowMode } from './modes/ExecuteWorkflowMode';
import { ManageFileMode } from './modes/ManageFileMode';
import { ManageImageMode } from './modes/ManageImageMode';
import { ManageAttachmentMode } from './modes/ManageAttachmentMode';
import { StubMode } from './modes/StubMode';
import { findTable } from './mock/metadata';
import { _setTableLookup } from './editors/filter/filterTree';
import { HostProvider, useHostSession } from './host/HostContext';
import { SettingsProvider } from './host/usePersistedSettings';
import { OpenSettingsProvider } from './host/useOpenSettings';

// Wire metadata lookup into the filter-tree encoder (avoids a circular import).
_setTableLookup(findTable);

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
        <ModeRouter modeId={activeId} themeMode={themeMode} setActiveId={setActiveId} />
      </SaveContextRoot>
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
    case 'exec-action':
      return <ExecuteActionMode themeMode={themeMode} category="oob" />;
    case 'exec-customapi':
      return <ExecuteActionMode themeMode={themeMode} category="custom-api" />;
    case 'exec-customaction':
      return <ExecuteActionMode themeMode={themeMode} category="custom-action" />;
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
