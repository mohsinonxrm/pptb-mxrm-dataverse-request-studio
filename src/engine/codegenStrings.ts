// String-literal escaping shared by the C# and PowerShell generators
// (codeGenerators.ts and sdkCodegen.ts).

/**
 * Escape text for a regular C# string literal ("…"): backslash first, then
 * the quote and control characters, so a header value like `C:\\x"y` can't
 * end the literal or change its meaning (#244).
 */
export function csharpString(s: string): string {
  return (
    s
      .replace(/\\/g, '\\\\')
      .replace(/"/g, '\\"')
      .replace(/\r/g, '\\r')
      .replace(/\n/g, '\\n')
      .replace(/\t/g, '\\t')
      // Other control characters, as \uXXXX like JSON, and the line
      // separators C# also refuses inside a regular literal (#232).
      .replace(
        // eslint-disable-next-line no-control-regex
        /[\u0000-\u001f\u0085\u2028\u2029]/g,
        (c) => `\\u${c.charCodeAt(0).toString(16).padStart(4, '0')}`,
      )
  );
}

/** Escape a JSON string for inclusion inside a C# verbatim string literal (@"…"). */
export function csharpVerbatimEscape(s: string): string {
  // In an @"…" verbatim literal, the only character that needs escaping is `"` → `""`.
  return s.replace(/"/g, '""');
}

/**
 * PowerShell verbatim (single-quoted) string: `it's` → `'it''s'`. PowerShell
 * also reads the curly quotes U+2018–U+201B as single quotes, so they're
 * doubled the same way, or a curly quote in a value would end the literal (#232).
 */
export function psQuote(v: string): string {
  return `'${v.replace(/['\u2018\u2019\u201a\u201b]/g, (q) => q + q)}'`;
}
