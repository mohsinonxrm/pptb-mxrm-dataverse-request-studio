// Cross-mode navigation with a pre-filled request (#3).
//
// A result card in one mode (e.g. WriteResultCard after a Create) can open
// another mode with part of its request already filled in: "Retrieve this
// record" switches to Retrieve Single with the table + record id set and
// runs it.
//
// Two pieces, both module-level so no provider is needed:
//   • the navigator — FrameHeader owns the mode picker, so it registers its
//     `setActiveId` via `useRegisterModeNavigator`;
//   • a pending-handoff slot per mode — `openInMode` fills it, then switches
//     mode; the receiving mode drains it once on mount (`takeModeHandoff`).
//
// Session memory only, never persisted. A handoff nobody drains is
// overwritten by the next one for the same mode.

import { useEffect } from 'react';

/** Payload each receiving mode understands. Add a key per mode that accepts handoffs. */
export interface ModeHandoffs {
  'retrieve-single': {
    table: string;
    recordId: string;
    /** Run the request as soon as the table's metadata is ready. */
    autoExecute?: boolean;
  };
}

type Navigator = (modeId: string) => void;

let modeNavigator: Navigator | null = null;
const pending = new Map<keyof ModeHandoffs, unknown>();

/** FrameHeader: register the mode picker's setter as the app-wide navigator. */
export function useRegisterModeNavigator(fn: Navigator): void {
  useEffect(() => {
    modeNavigator = fn;
    return () => {
      if (modeNavigator === fn) modeNavigator = null;
    };
  }, [fn]);
}

/** True when a navigator is mounted (i.e. `openInMode` will switch modes). */
export function canOpenInMode(): boolean {
  return modeNavigator !== null;
}

/** Switch to `modeId` with `payload` waiting for it. */
export function openInMode<K extends keyof ModeHandoffs>(
  modeId: K,
  payload: ModeHandoffs[K],
): void {
  if (!modeNavigator) return;
  pending.set(modeId, payload);
  modeNavigator(modeId);
}

/** Receiving mode, on mount: take (and clear) the payload waiting for it, if any. */
export function takeModeHandoff<K extends keyof ModeHandoffs>(
  modeId: K,
): ModeHandoffs[K] | undefined {
  const payload = pending.get(modeId) as ModeHandoffs[K] | undefined;
  pending.delete(modeId);
  return payload;
}
