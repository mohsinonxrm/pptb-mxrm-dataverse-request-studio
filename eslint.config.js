// ESLint (#273), shared with FetchXML Studio and Command Bar Studio:
// recommended JS + TypeScript rules, React hooks and refresh, jsx-a11y.
import js from '@eslint/js';
import globals from 'globals';
import reactHooks from 'eslint-plugin-react-hooks';
import reactRefresh from 'eslint-plugin-react-refresh';
import jsxA11y from 'eslint-plugin-jsx-a11y';
import tseslint from 'typescript-eslint';

export default tseslint.config(
  { ignores: ['dist', 'coverage', 'node_modules', '.claude'] },
  {
    extends: [js.configs.recommended, ...tseslint.configs.recommended],
    files: ['**/*.{ts,tsx}'],
    languageOptions: {
      ecmaVersion: 2022,
      globals: globals.browser,
    },
    plugins: {
      'react-hooks': reactHooks,
      'react-refresh': reactRefresh,
      'jsx-a11y': jsxA11y,
    },
    rules: {
      ...reactHooks.configs.recommended.rules,
      // Off here (FetchXML / Command Bar Studio warn): DRS keeps a component's
      // helpers and types next to it (`emptyApply`, `targetsTitle`, …), which
      // only costs a full reload instead of a hot one in `npm run dev`.
      'react-refresh/only-export-components': 'off',
      // `_`-prefixed names are deliberately unused here too: `const { dirty: _dirty, ...rest } = s`
      // drops a field before saving.
      '@typescript-eslint/no-unused-vars': [
        'error',
        {
          argsIgnorePattern: '^_',
          varsIgnorePattern: '^_',
          destructuredArrayIgnorePattern: '^_',
          ignoreRestSiblings: true,
        },
      ],
      '@typescript-eslint/no-explicit-any': 'error',
      // Copy and paste go through the ToolBox bridge, which reports failures
      // (#242): `copyText` in shell/copy.tsx, `readClipboardText` in the bridge.
      'no-restricted-properties': [
        'error',
        {
          object: 'navigator',
          property: 'clipboard',
          message: 'Use copyText (shell/copy.tsx) or readClipboardText (host/pptbBridge.ts).',
        },
      ],
    },
  },
  {
    // The bridge is where the browser clipboard is used, and tests stub it.
    files: ['src/host/pptbBridge.ts', '**/*.test.{ts,tsx}'],
    rules: { 'no-restricted-properties': 'off' },
  },
  {
    // Node-side config files (Vite, Vitest, ESLint itself) and CI scripts.
    files: ['*.config.{js,ts}', '.github/scripts/**/*.mjs'],
    languageOptions: {
      globals: globals.node,
    },
  },
  {
    files: ['**/*.test.{ts,tsx}', 'src/test/**'],
    rules: {
      '@typescript-eslint/no-explicit-any': 'off',
    },
  },
);
