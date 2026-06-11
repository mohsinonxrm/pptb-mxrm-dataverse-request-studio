// useOpenSettings — app-wide context that lets any component trigger the
// Settings drawer without prop-drilling the open-state callback through
// every mode → editor → primitive call chain.
//
// Pattern:
//   1. Wrap the root tree in <OpenSettingsProvider>.
//   2. FrameHeader calls `useRegisterOpenSettings(() => setSettingsOpen(true))`
//      once it has its local open-state setter.
//   3. Any deep component (e.g. TargetEditor) calls `useOpenSettings()` to
//      get a stable `() => void` callback that opens the drawer.
//
// The registration uses a ref so the provider re-renders zero times when
// FrameHeader mounts/unmounts.

import { createContext, useCallback, useContext, useEffect, useRef, type ReactNode } from 'react';

type OpenFn = () => void;

interface OpenSettingsCtx {
  open: OpenFn;
  register: (fn: OpenFn) => void;
}

const Ctx = createContext<OpenSettingsCtx | null>(null);

export function OpenSettingsProvider({ children }: { children: ReactNode }) {
  const ref = useRef<OpenFn | null>(null);

  const open = useCallback(() => ref.current?.(), []);
  const register = useCallback((fn: OpenFn) => {
    ref.current = fn;
  }, []);

  return <Ctx.Provider value={{ open, register }}>{children}</Ctx.Provider>;
}

/** Call this in FrameHeader to register its local `setSettingsOpen(true)` callback. */
export function useRegisterOpenSettings(fn: OpenFn) {
  const ctx = useContext(Ctx);
  useEffect(() => {
    ctx?.register(fn);
  }, [ctx, fn]);
}

/**
 * Returns a stable `() => void` that opens the Settings drawer.
 * Returns a no-op when called outside the provider (safe fallback).
 */
export function useOpenSettings(): OpenFn {
  const ctx = useContext(Ctx);
  return ctx?.open ?? (() => {});
}
