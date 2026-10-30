// useOpenScope — app-wide context that lets any component open the header's
// scope chip popover (#42) without prop-drilling the open-state callback
// through every mode → editor → primitive call chain.
//
// Pattern:
//   1. Wrap the root tree in <OpenScopeProvider>.
//   2. ScopeChip calls `useRegisterOpenScope(() => setOpen(true))` once it
//      has its local open-state setter.
//   3. Any deep component (e.g. TargetEditor's empty state) calls
//      `useOpenScope()` to get a stable `() => void` callback that opens it.
//
// The registration uses a ref so the provider re-renders zero times when
// the chip mounts/unmounts.

import { createContext, useCallback, useContext, useEffect, useRef, type ReactNode } from 'react';

type OpenFn = () => void;

interface OpenScopeCtx {
  open: OpenFn;
  register: (fn: OpenFn) => void;
}

const Ctx = createContext<OpenScopeCtx | null>(null);

export function OpenScopeProvider({ children }: { children: ReactNode }) {
  const ref = useRef<OpenFn | null>(null);

  const open = useCallback(() => ref.current?.(), []);
  const register = useCallback((fn: OpenFn) => {
    ref.current = fn;
  }, []);

  return <Ctx.Provider value={{ open, register }}>{children}</Ctx.Provider>;
}

/** Call this in ScopeChip to register its local `setOpen(true)` callback. */
export function useRegisterOpenScope(fn: OpenFn) {
  const ctx = useContext(Ctx);
  useEffect(() => {
    ctx?.register(fn);
  }, [ctx, fn]);
}

/**
 * Returns a stable `() => void` that opens the scope chip's popover.
 * Returns a no-op when called outside the provider (safe fallback).
 */
export function useOpenScope(): OpenFn {
  const ctx = useContext(Ctx);
  return ctx?.open ?? (() => {});
}
