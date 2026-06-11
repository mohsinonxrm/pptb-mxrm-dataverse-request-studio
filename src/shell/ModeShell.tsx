import type { ReactNode } from 'react';
import { useStudioStyles } from '../primitives/styles';

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
  const s = useStudioStyles();
  return (
    <>
      <div className={s.body}>
        {sidebar}
        {children}
      </div>
      {urlBar}
    </>
  );
}
