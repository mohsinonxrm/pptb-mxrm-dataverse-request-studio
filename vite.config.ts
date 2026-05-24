import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// Vite config — mirrors FetchXML Studio's Monaco-bundling strategy.
//
//   • `optimizeDeps.include: ['monaco-editor']`
//     Pre-bundle Monaco at dev-time so Vite doesn't try to transform its
//     deep internal imports lazily (avoids CJS/ESM interop edge cases that
//     surface as runtime "module not found" errors).
//
//   • `rollupOptions.output.manualChunks`
//     Keep Monaco's huge language packs out of the main bundle. The split
//     is loaded on-demand when the Code tab opens for the first time —
//     keeps the boot bundle lean.
//
// Why this combination matters for PPTB: the alternative (Monaco's default
// AMD loader from `cdn.jsdelivr.net`) is blocked by PPTB's CSP. Bundling
// locally + calling `loader.config({ monaco })` in CodeView.tsx routes
// Monaco entirely through the same-origin chunks the iframe is allowed
// to load.
export default defineConfig({
  plugins: [react()],
  server: { port: 5173, open: true },
  optimizeDeps: {
    include: ['monaco-editor'],
  },
  build: {
    target: 'es2020',
    // Source maps disabled for production builds. They add ~17 MB to the
    // npm package (Monaco worker .map files are huge) and PPTB users
    // don't debug the production bundle — development uses `npm run dev`
    // which provides full HMR + source maps. Flip to `true` or `'hidden'`
    // locally if you need to debug a specific production-only issue.
    sourcemap: false,
    rollupOptions: {
      output: {
        manualChunks: {
          'monaco-editor': ['monaco-editor'],
        },
      },
    },
  },
});
