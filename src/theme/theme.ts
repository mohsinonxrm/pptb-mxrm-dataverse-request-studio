import { webLightTheme, webDarkTheme, type Theme } from '@fluentui/react-components';

// Note: custom keys won't surface via griffel's `tokens.*` proxy (build-time
// const). We keep our 11 brand/method colors as CSS variables on
// :root[data-theme] (declared in index.html) and read them via
// `style={{ color: 'var(--color-request-read)' }}`.
//
// We still spread the v9 themes so we can extend later without changing call
// sites in components.
export const studioLight: Theme = { ...webLightTheme };
export const studioDark: Theme = { ...webDarkTheme };

export type ThemeMode = 'light' | 'dark';

export function applyThemeAttr(mode: ThemeMode): void {
  document.documentElement.setAttribute('data-theme', mode);
}

// Group → CSS-var name. Used everywhere a "request group color" is needed.
export const groupColorVar = (group: string): string =>
  `var(--color-request-${group})`;

export const methodColorVar = (method: string): string =>
  `var(--color-method-${method.toLowerCase()})`;
