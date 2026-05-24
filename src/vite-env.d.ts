// Vite's bundler-specific ambient types — required for the `?worker` /
// `?url` / `?raw` query-string imports used in CodeView.tsx (Monaco workers)
// and anywhere we adopt `import x from 'foo?worker'` going forward.
//
// Without this reference, `tsc` reports
//   TS2307: Cannot find module 'monaco-editor/.../editor.worker?worker'
// even though Vite resolves it correctly at bundle time.
/// <reference types="vite/client" />
