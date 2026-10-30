import { useCallback, useLayoutEffect, useRef } from 'react';

/**
 * A function that keeps the same identity across renders but always runs the
 * latest `fn`. For handlers handed to something that holds on to them, like
 * the Save / Saved-library buttons in the header: a mode publishes its save
 * context only when its data changes, so a plain closure would run with the
 * data of that render (the previous org's table list, a stale request).
 */
export function useLatestCallback<A extends unknown[], R>(
  fn: (...args: A) => R,
): (...args: A) => R {
  const latest = useRef(fn);
  useLayoutEffect(() => {
    latest.current = fn;
  });
  return useCallback((...args: A) => latest.current(...args), []);
}
