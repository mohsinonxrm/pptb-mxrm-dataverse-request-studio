/**
 * Up to two initials from a record name: "Contoso Ltd" → "CL", "(Test)
 * Account" → "TA", "" → "?". Each word gives its first letter or digit
 * (with any accents on it); words with neither, like an emoji, are skipped.
 */
export function recordInitials(name: string): string {
  const letters = name
    .normalize('NFC')
    .split(/\s+/)
    .map((word) => /[\p{L}\p{N}]\p{M}*/u.exec(word)?.[0])
    .filter((c): c is string => !!c)
    // Keep a letter whose capital is longer ("ß" → "SS") as it is.
    .map((c) => (c.toUpperCase().length === c.length ? c.toUpperCase() : c));
  if (letters.length === 0) return '?';
  return letters.length === 1 ? letters[0] : letters[0] + letters[letters.length - 1];
}
