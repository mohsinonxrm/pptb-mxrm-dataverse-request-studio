/// <reference types="vitest/config" />
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { gzipSync } from 'node:zlib';
import { defineConfig, type Plugin } from 'vite';
import react from '@vitejs/plugin-react';

// Bundle the tool icon at dist/icons/icon.svg, where package.json's top-level
// `icon` points. PPTB resolves `icon` relative to the dist root, so the
// repo-root `icons/` folder is never seen by the host (#25). Prerelease
// versions (e.g. 1.1.4-beta.1) get the gold Insider variant; releases get the
// regular icon. https://docs.powerplatformtoolbox.com/tool-development/publishing
function toolIcon(): Plugin {
  return {
    name: 'drs-tool-icon',
    apply: 'build',
    generateBundle() {
      // Resolve against this config file, not the process cwd.
      const fromRoot = (rel: string) => fileURLToPath(new URL(rel, import.meta.url));
      const { version } = JSON.parse(readFileSync(fromRoot('package.json'), 'utf8')) as {
        version: string;
      };
      const source = version.includes('-') ? 'icons/icon-insider.svg' : 'icons/icon.svg';
      this.emitFile({
        type: 'asset',
        fileName: 'icons/icon.svg',
        source: readFileSync(fromRoot(source), 'utf8'),
      });
    },
  };
}

// Opt-in bundle report (#22): `npm run build:analyze` (or ANALYZE=1).
// Dependency-free: prints every emitted JS/CSS chunk with raw + gzip size
// and the heaviest packages inside each chunk (by rendered bytes after tree
// shaking), so a regression like Monaco landing in the boot path shows up.
// Registered always but inert unless the mode is `analyze` or ANALYZE is set,
// so the normal build is unchanged.
function bundleReport(): Plugin {
  let enabled = false;
  const kb = (n: number) => `${(n / 1024).toFixed(1)} kB`.padStart(11);
  const pkgOf = (id: string) => {
    const path = id.replace(/\\/g, '/');
    const m = path.match(/node_modules\/((?:@[^/]+\/)?[^/]+)/);
    return m ? m[1] : path.includes('/src/') ? '(app src)' : '(other)';
  };
  return {
    name: 'drs-bundle-report',
    apply: 'build',
    configResolved(config) {
      enabled = config.mode === 'analyze' || !!process.env.ANALYZE;
    },
    generateBundle(_opts, bundle) {
      if (!enabled) return;
      const rows = Object.values(bundle)
        .filter((f) => /\.(js|css)$/.test(f.fileName))
        .map((f) => {
          const code = f.type === 'chunk' ? f.code : f.source;
          const raw = typeof code === 'string' ? Buffer.byteLength(code) : code.byteLength;
          return { f, raw, gzip: gzipSync(code).byteLength };
        })
        .sort((a, b) => b.raw - a.raw);
      // Boot path = entry chunk(s) + everything they import statically.
      const boot = new Set<string>();
      const visit = (name: string) => {
        const c = bundle[name];
        if (!c || c.type !== 'chunk' || boot.has(name)) return;
        boot.add(name);
        c.imports.forEach(visit);
        c.viteMetadata?.importedCss.forEach((css) => boot.add(css));
      };
      Object.values(bundle).forEach((f) => f.type === 'chunk' && f.isEntry && visit(f.fileName));

      // Heaviest entries of a size map, as indented report lines.
      const top = (sizes: Map<string, number>, n: number) =>
        [...sizes.entries()]
          .sort((a, b) => b[1] - a[1])
          .slice(0, n)
          .map(([name, size]) => `${kb(size)}              ${name}`);

      const lines = ['', 'Bundle report — raw / gzip; boot = loaded at startup', ''];
      const small = { count: 0, raw: 0, gzip: 0 };
      const appFiles = new Map<string, number>();
      for (const { f, raw, gzip } of rows) {
        const kind = boot.has(f.fileName)
          ? 'boot'
          : f.type === 'chunk'
            ? 'lazy'
            : f.fileName.endsWith('.js')
              ? 'worker'
              : 'css';
        if (f.type === 'chunk') {
          for (const [id, m] of Object.entries(f.modules)) {
            if (pkgOf(id) === '(app src)')
              appFiles.set(id.replace(/\\/g, '/').replace(/^.*\/src\//, 'src/'), m.renderedLength);
          }
        }
        // Monaco's ~80 per-language chunks are each a few kB — summarise them.
        if (raw < 50 * 1024 && !boot.has(f.fileName)) {
          small.count++;
          small.raw += raw;
          small.gzip += gzip;
          continue;
        }
        lines.push(`${kb(raw)} ${kb(gzip)}  ${kind.padEnd(6)} ${f.fileName}`);
        if (f.type !== 'chunk' || raw < 100 * 1024) continue;
        // Per-package bytes are pre-minification (rollup renderedLength).
        const byPkg = new Map<string, number>();
        for (const [id, m] of Object.entries(f.modules)) {
          byPkg.set(pkgOf(id), (byPkg.get(pkgOf(id)) ?? 0) + m.renderedLength);
        }
        lines.push(...top(byPkg, 10));
      }
      lines.push(
        `${kb(small.raw)} ${kb(small.gzip)}  lazy   (${small.count} chunks under 50 kB)`,
        '',
        'App source by folder (pre-minification):',
        ...top(
          [...appFiles].reduce((m, [file, n]) => {
            const dir = file.split('/').slice(0, 2).join('/');
            return m.set(dir, (m.get(dir) ?? 0) + n);
          }, new Map<string, number>()),
          10,
        ),
        '',
        'Largest app source files (pre-minification):',
        ...top(appFiles, 10),
      );
      const bootRows = rows.filter(({ f }) => boot.has(f.fileName));
      const sum = (k: 'raw' | 'gzip') => bootRows.reduce((n, r) => n + r[k], 0);
      lines.push('', `Boot path (${bootRows.length} files): ${kb(sum('raw'))} ${kb(sum('gzip'))}`);
      console.log(lines.join('\n'));
    },
  };
}

// Vite config — mirrors FetchXML Studio's Monaco-bundling strategy.
//
//   • `optimizeDeps.include: ['monaco-editor']`
//     Pre-bundle Monaco at dev-time so Vite doesn't try to transform its
//     deep internal imports lazily (avoids CJS/ESM interop edge cases that
//     surface as runtime "module not found" errors).
//
//   • Lazy Monaco (src/monaco/MonacoEditor.tsx, React.lazy)
//     Keep Monaco and its huge language packs out of the main bundle. The
//     split is loaded on demand the first time any editor renders — keeps
//     the boot bundle lean.
//
// Why this combination matters for PPTB: the alternative (Monaco's default
// AMD loader from `cdn.jsdelivr.net`) is blocked by PPTB's CSP. Bundling
// locally + calling `loader.config({ monaco })` in src/monaco/monacoSetup.ts routes
// Monaco entirely through the same-origin chunks the iframe is allowed
// to load.
export default defineConfig({
  // Relative asset URLs (#212). PPTB Desktop serves the tool at the root of
  // its own origin, but the PPTB VS Code extension serves `dist/` from a
  // nested webview resource URL: a root-absolute `/assets/…` built at runtime
  // (lazy chunks and their preloads, Monaco's workers) or written in CSS
  // (the codicon font) resolves outside the tool and fails. CI checks the
  // build for root-absolute asset URLs.
  base: './',
  plugins: [react(), toolIcon(), bundleReport()],
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
    // No `manualChunks` for Monaco (#22). The old
    // `{ 'monaco-editor': ['monaco-editor'] }` also pulled Vite's shared
    // preload helper into the Monaco chunk, so the entry imported that chunk
    // statically and all of Monaco (~930 kB gzip) loaded at startup. Monaco
    // is now reached only through a dynamic import (src/monaco/MonacoEditor),
    // which splits it — and its per-language chunks — out on its own.
  },
  test: {
    // jsdom provides DOMParser, which the CSDL parser (host/csdlProvider) relies on.
    environment: 'jsdom',
    include: ['src/**/*.test.{ts,tsx}'],
  },
});
