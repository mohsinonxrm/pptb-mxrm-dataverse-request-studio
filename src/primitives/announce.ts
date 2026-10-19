// Screen-reader announcements for async outcomes (#20, WCAG 4.1.3 Status
// Messages). Execute runs in the background and its result lands in another
// tab, so without a live region a screen-reader user gets no feedback that
// the request ran, failed or how many rows came back.
//
// One visually hidden polite live region, appended to <body> on first use.
// It is deliberately separate from Fluent's AriaLiveAnnouncer: that one
// would also make every MessageBar announce itself on mount.

import { useEffect } from 'react';
import type { ExecResult } from '../engine/dataverseExecutor';

let region: HTMLElement | null = null;
let timer: number | undefined;

function liveRegion(): { el: HTMLElement; fresh: boolean } {
  if (region && document.body.contains(region)) return { el: region, fresh: false };
  const el = document.createElement('div');
  el.setAttribute('role', 'status');
  el.setAttribute('aria-live', 'polite');
  el.setAttribute('aria-atomic', 'true');
  el.setAttribute('data-drs-announcer', '');
  Object.assign(el.style, {
    position: 'absolute',
    width: '1px',
    height: '1px',
    margin: '-1px',
    padding: '0',
    border: '0',
    overflow: 'hidden',
    clip: 'rect(0 0 0 0)',
    whiteSpace: 'nowrap',
  });
  document.body.appendChild(el);
  region = el;
  return { el, fresh: true };
}

/** Announce `message` politely. Repeating the same text is announced again. */
export function announce(message: string): void {
  if (typeof document === 'undefined') return;
  const { el, fresh } = liveRegion();
  el.textContent = '';
  window.clearTimeout(timer);
  // A region inserted in the same tick as its text is often not read, and
  // clearing first lets an identical message be read twice.
  timer = window.setTimeout(() => (el.textContent = message), fresh ? 250 : 50);
}

/** One-line spoken summary of an Execute result. */
export function describeResult(result: ExecResult): string {
  const status = result.status
    ? `${result.status}${result.statusText ? ` ${result.statusText}` : ''}`
    : 'no HTTP status';
  if (!result.ok) {
    const err = (result.body as { error?: { message?: string } } | null)?.error;
    return `Request failed, ${status}.${err?.message ? ` ${err.message}` : ''}`;
  }
  const env = result.body as { value?: unknown[] } | null;
  const rows = Array.isArray(env?.value)
    ? `, ${env!.value!.length} row${env!.value!.length === 1 ? '' : 's'}`
    : '';
  return `Request completed, ${status}${rows}, ${result.ms} ms.`;
}

// Results already announced — switching back to the Results tab remounts
// the view with the same result object, which shouldn't be read again.
const announced = new WeakSet<ExecResult>();

/** Announce each new Execute result once. */
export function useAnnounceResult(result: ExecResult | null): void {
  useEffect(() => {
    if (!result || announced.has(result)) return;
    announced.add(result);
    announce(describeResult(result));
  }, [result]);
}
