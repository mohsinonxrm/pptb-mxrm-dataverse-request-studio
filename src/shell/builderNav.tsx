// Builder navigation for an advisory's Open (#294): show the Builder tab,
// then focus the first field of the pane the mode just selected.
//
// ModeShell provides it, MainTabs registers how to show the Builder, and the
// URL bar's advisory drawer uses it, so every mode gets the same behaviour
// without passing its tab setter around.

import { createContext, useContext, useEffect, useMemo, useRef, type ReactNode } from 'react';

interface BuilderNav {
  /** MainTabs: how to show the Builder tab. Returns the unregister. */
  register: (showBuilder: () => void) => () => void;
  /** Show the Builder tab (a no-op when no MainTabs registered). */
  showBuilder: () => void;
}

const BuilderNavContext = createContext<BuilderNav | null>(null);

export function BuilderNavProvider({ children }: { children: ReactNode }) {
  const current = useRef<(() => void) | null>(null);
  const value = useMemo<BuilderNav>(
    () => ({
      register: (fn) => {
        current.current = fn;
        return () => {
          if (current.current === fn) current.current = null;
        };
      },
      showBuilder: () => current.current?.(),
    }),
    [],
  );
  return <BuilderNavContext.Provider value={value}>{children}</BuilderNavContext.Provider>;
}

/** MainTabs: register how to show the Builder tab. */
export function useRegisterShowBuilder(showBuilder: () => void) {
  const nav = useContext(BuilderNavContext);
  const latest = useRef(showBuilder);
  latest.current = showBuilder;
  useEffect(() => nav?.register(() => latest.current()), [nav]);
}

/** The attribute MainTabs puts on the pane body, so Open can find its fields. */
const BUILDER_BODY_ATTR = 'data-builder-body';

const FIELD_SELECTOR = [
  'input:not([disabled]):not([type="hidden"])',
  'textarea:not([disabled])',
  'select:not([disabled])',
  '[role="combobox"]:not([aria-disabled="true"])',
  '[contenteditable="true"]',
].join(', ');

/** A pane marks the element of a node that shares the pane with others (a
 *  required parameter, say) so Open lands on it rather than the pane's top. */
const FOCUS_NODE_ATTR = 'data-focus-node';

/** How long Open waits for a node's element: a pane may render its fields
 *  only once metadata arrives. About half a second at 60 fps. */
const MAX_FRAMES = 30;

/** Anything else focusable, for a marked node whose input is a button. */
const FOCUSABLE_SELECTOR = 'button:not([disabled]), [tabindex]:not([tabindex="-1"])';

/** Only the latest Open moves focus; an older one still waiting gives up. */
let openSeq = 0;

/**
 * Once the pane has rendered, focus the first field of `nodeId`'s element
 * when the pane marks one, otherwise the pane's first field, otherwise the
 * pane body itself (made focusable just for this).
 */
export function focusBuilderField(nodeId?: string, doc: Document = document) {
  const seq = ++openSeq;
  const before = doc.activeElement;
  let frames = 0;
  const attempt = () => {
    if (seq !== openSeq) return;
    // The user moved on to another field meanwhile: leave their focus alone.
    const now = doc.activeElement;
    if (now !== before && now instanceof HTMLElement && now.matches(FIELD_SELECTOR)) return;
    const body = doc.querySelector<HTMLElement>(`[${BUILDER_BODY_ATTR}]`);
    const marked = nodeId
      ? body?.querySelector<HTMLElement>(`[${FOCUS_NODE_ATTR}="${CSS.escape(nodeId)}"]`)
      : null;
    const field = marked?.querySelector<HTMLElement>(FIELD_SELECTOR);
    // The mode's state update renders the pane in the first frame. A marked
    // node may get its fields a little later, once metadata arrives.
    if (frames++ < 2 || (marked && !field && frames < MAX_FRAMES)) {
      requestAnimationFrame(attempt);
      return;
    }
    if (!body) return;
    const target =
      field ??
      marked?.querySelector<HTMLElement>(FOCUSABLE_SELECTOR) ??
      body.querySelector<HTMLElement>(FIELD_SELECTOR) ??
      focusableBody(body);
    target.focus();
    target.scrollIntoView?.({ block: 'nearest' });
  };
  requestAnimationFrame(attempt);
}

/** The pane body, focusable until focus leaves it. It isn't permanently, so
 *  the browser still treats it as a keyboard-scrollable region. */
function focusableBody(body: HTMLElement): HTMLElement {
  if (!body.hasAttribute('tabindex')) {
    body.setAttribute('tabindex', '-1');
    body.addEventListener('blur', () => body.removeAttribute('tabindex'), { once: true });
  }
  return body;
}

/** The drawer's Open: show the Builder, select the node, focus its field. */
export function useOpenAdvisoryNode(onSelectNode?: (nodeId: string) => void) {
  const nav = useContext(BuilderNavContext);
  if (!onSelectNode) return undefined;
  return (nodeId: string) => {
    nav?.showBuilder();
    onSelectNode(nodeId);
    focusBuilderField(nodeId);
  };
}
