// ── Stale-chunk detection (#187) ──────────────────────────────────────
//
// Monaco and the less-used modes are code-split (#22). If DRS is updated (or
// rebuilt) while a tab is open, the old page still asks for the old hashed
// chunk names, which no longer exist, so the dynamic import fails. Reloading
// the pane can't fix that (React.lazy caches the rejection and the file is
// gone); only reopening the tool loads the new build. The error boundaries
// use this to show that specific advice instead of the generic fallback.

export const STALE_BUILD_MESSAGE =
  'Dataverse Request Studio was updated — close and reopen the tool to load the new version.';

const CHUNK_ERROR_PATTERNS = [
  /Failed to fetch dynamically imported module/i,
  /error loading dynamically imported module/i, // Firefox wording
  /Importing a module script failed/i, // Safari wording
  /Unable to preload CSS/i, // Vite's CSS preload failure
  /Loading (CSS )?chunk [\w-]+ failed/i, // webpack-style ChunkLoadError
];

let preloadFailed = false;

/** Vite fires `vite:preloadError` when a chunk (or its CSS) fails to load. */
export function installChunkLoadErrorListener() {
  if (typeof window === 'undefined') return;
  window.addEventListener('vite:preloadError', () => {
    // Don't preventDefault: the import should still reject so the nearest
    // error boundary renders. We only remember that it was a chunk failure.
    preloadFailed = true;
  });
}

export function isChunkLoadError(error: unknown): boolean {
  if (!error) return false;
  const e = error as { name?: unknown; message?: unknown };
  if (e.name === 'ChunkLoadError') return true;
  const message = typeof e.message === 'string' ? e.message : String(error);
  if (CHUNK_ERROR_PATTERNS.some((re) => re.test(message))) return true;
  // A preload failure was reported and this error is the import rejecting.
  return preloadFailed && e.name === 'TypeError';
}
