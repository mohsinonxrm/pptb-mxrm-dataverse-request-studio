import type { ReactNode } from 'react';
import { useClearAnnouncementOnModeSwitch } from '../primitives/announce';
import { useStudioStyles } from '../primitives/styles';
import { RecordPickerMemoryScope } from '../primitives/RecordPicker';
import { BuilderNavProvider } from './builderNav';

/**
 * Wraps a mode's three regions in the correct DOM structure:
 *
 *   <body>  (grid: 300px sidebar / 1fr main)
 *     <Sidebar />
 *     <Main />
 *   </body>
 *   <UrlBar />     ← full-width, spans below both sidebar and main
 *
 * The parent Frame is `grid: auto 1fr auto`, so the body row stretches and the
 * URL bar sits flush at the bottom across the full window width.
 */
export function ModeShell({
  sidebar,
  children,
  urlBar,
}: {
  sidebar: ReactNode;
  /** Main pane content (typically MainTabs) */
  children: ReactNode;
  urlBar: ReactNode;
}) {
  // Each mode renders its own ModeShell, so this runs on every mode switch.
  useClearAnnouncementOnModeSwitch();
  const s = useStudioStyles();
  // The picker memory scope lives here because ModeShell outlives the
  // sidebar panes (#189): a pane switch keeps the record picker's mode.
  return (
    <RecordPickerMemoryScope>
      {/* An advisory's Open in the URL bar shows the Builder in MainTabs (#294). */}
      <BuilderNavProvider>
        <div className={s.body}>
          {sidebar}
          {children}
        </div>
        {urlBar}
      </BuilderNavProvider>
    </RecordPickerMemoryScope>
  );
}
