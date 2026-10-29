import { useState } from 'react';
import {
  Breadcrumb,
  BreadcrumbButton,
  BreadcrumbDivider,
  BreadcrumbItem,
  Button,
  Tooltip,
  ToolbarDivider,
  Menu,
  MenuTrigger,
  MenuPopover,
  MenuList,
  MenuItem,
  Badge,
  tokens,
} from '@fluentui/react-components';
import {
  WeatherSunny20Regular,
  WeatherMoon20Regular,
  Share20Regular,
  MoreHorizontal20Regular,
  Database20Regular,
  Settings20Regular,
} from '@fluentui/react-icons';
import { useStudioStyles } from '../primitives/styles';
import { RequestTypePicker } from '../primitives/RequestTypePicker';
import { SettingsDrawer } from './SettingsDrawer';
import { usePersistedSettings } from '../host/usePersistedSettings';
import { useRegisterOpenSettings } from '../host/useOpenSettings';
import { useAccessMode } from '../host/useAccessMode';
import type { ThemeMode } from '../theme/theme';
import { getEnv } from '../mock/environment';
import { useHostSession } from '../host/HostContext';
import { useSaveContext } from '../state/SaveContext';
import { useRegisterModeNavigator } from '../state/modeHandoff';
import { SaveButton, SavedLibraryButton } from '../editors/SavedRequestsDialog';
import { EnvironmentBadge } from './confirmWrite';

export function FrameHeader({
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
  // PPTB owns the env switcher, the tool's chrome, and the theme toggle. Hide
  // those affordances when embedded so we don't duplicate the host chrome — but
  // keep them in standalone mode (the local dev experience needs them).
  const host = useHostSession();
  const embedded = host.embedded;
  // The connection name, else the environment's host name.
  const connectionLabel =
    host.connectionName || host.environmentUrl.replace(/^https?:\/\//, '').replace(/\/.*$/, '');
  const env = getEnv();
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [settings, updateSettings] = usePersistedSettings();
  const { accessSummary } = useAccessMode();
  // Register the open-settings callback so any component deep in the tree
  // (e.g. TargetEditor's guidance banner) can open the drawer without prop-drilling.
  useRegisterOpenSettings(() => setSettingsOpen(true));
  // The mode picker lives here, so this is also where result cards switch
  // modes from (e.g. WriteResultCard's "Retrieve this record", #3).
  useRegisterModeNavigator(setActiveId);
  // The current mode (if any) publishes its save context into the
  // SaveContextRoot in App.tsx; we read it here so the top-right Save +
  // Saved-library buttons drive that mode's persistence. If no mode is
  // publishing (e.g. specialty modes with no persistable state), the
  // buttons hide so the header doesn't carry inert chrome.
  const saveCtx = useSaveContext();
  return (
    <header className={s.header}>
      {!embedded && (
        <>
          <span className={s.brand} style={{ flexShrink: 0 }}>
            <span className={s.brandDot} />
            <span style={{ whiteSpace: 'nowrap' }}>Dataverse Request Studio</span>
          </span>
          <span style={{ color: tokens.colorNeutralForeground3, fontSize: 12, flexShrink: 0 }}>
            ·
          </span>
        </>
      )}
      <div style={{ display: 'flex', minWidth: 0, overflow: 'hidden', flexShrink: 1 }}>
        <Breadcrumb size="small">
          {!embedded && (
            <>
              <BreadcrumbItem>
                <BreadcrumbButton>{env.name}</BreadcrumbButton>
              </BreadcrumbItem>
              <BreadcrumbDivider />
            </>
          )}
          <BreadcrumbItem>
            <RequestTypePicker activeId={activeId} onPick={setActiveId} />
          </BreadcrumbItem>
        </Breadcrumb>
      </div>
      <div className={s.spacer} />
      {/* The connection the requests go to, also when embedded, with
          Production highlighted (#241). */}
      {embedded && host.connected && (connectionLabel || host.environment) && (
        <Tooltip
          content={`Requests go to ${host.connectionName || 'this connection'}${
            host.environmentUrl ? ` (${host.environmentUrl.replace(/\/$/, '')})` : ''
          }`}
          relationship="description"
        >
          <span
            data-testid="environment-indicator"
            style={{
              display: 'inline-flex',
              alignItems: 'center',
              gap: 6,
              flexShrink: 1,
              minWidth: 0,
              maxWidth: 260,
              marginRight: 8,
            }}
          >
            {connectionLabel && (
              <span
                style={{
                  fontSize: tokens.fontSizeBase200,
                  color: tokens.colorNeutralForeground2,
                  whiteSpace: 'nowrap',
                  overflow: 'hidden',
                  textOverflow: 'ellipsis',
                }}
              >
                {connectionLabel}
              </span>
            )}
            <EnvironmentBadge environment={host.environment} />
          </span>
        </Tooltip>
      )}
      {!embedded && (
        <>
          <Badge
            appearance="ghost"
            icon={<Database20Regular />}
            style={{
              flexShrink: 0,
              maxWidth: 280,
              overflow: 'hidden',
              whiteSpace: 'nowrap',
              textOverflow: 'ellipsis',
            }}
          >
            {env.host}
          </Badge>
          <ToolbarDivider />
          <Tooltip
            content={themeMode === 'dark' ? 'Switch to light' : 'Switch to dark'}
            relationship="label"
          >
            <Button
              icon={themeMode === 'dark' ? <WeatherSunny20Regular /> : <WeatherMoon20Regular />}
              appearance="subtle"
              onClick={() => setThemeMode(themeMode === 'dark' ? 'light' : 'dark')}
            />
          </Tooltip>
        </>
      )}
      {/* Save + Saved-library buttons — driven by the current mode's
          SaveContext. Hidden when no mode is publishing (e.g. legacy
          modes without persistable state). The Save button is dirty-
          aware (filled icon when there are unsaved changes, subtle
          when the in-memory state already matches a saved entry). */}
      {saveCtx && (
        <>
          <SavedLibraryButton
            modeId={saveCtx.modeId}
            onLoad={saveCtx.onLoadSaved}
            currentId={saveCtx.lastSavedId}
            dirty={saveCtx.dirty}
          />
          <SaveButton
            state={saveCtx.state}
            modeId={saveCtx.modeId}
            dirty={saveCtx.dirty}
            lastSavedId={saveCtx.lastSavedId}
            onSaved={saveCtx.onSaved}
          />
        </>
      )}
      {!embedded && (
        <Tooltip content="Share" relationship="label">
          <Button icon={<Share20Regular />} appearance="subtle" />
        </Tooltip>
      )}
      <Menu>
        <MenuTrigger disableButtonEnhancement>
          <Button
            icon={<MoreHorizontal20Regular />}
            appearance="subtle"
            aria-label="More actions"
            title="More actions"
          />
        </MenuTrigger>
        <MenuPopover>
          <MenuList>
            <MenuItem icon={<Settings20Regular />} onClick={() => setSettingsOpen(true)}>
              Settings
            </MenuItem>
          </MenuList>
        </MenuPopover>
      </Menu>
      <SettingsDrawer
        open={settingsOpen}
        onClose={() => setSettingsOpen(false)}
        settings={settings}
        onSettingsChange={updateSettings}
        accessSummary={accessSummary}
      />
    </header>
  );
}
