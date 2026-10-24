// Response sizes in bytes, as the "KB" figures in the results show them.
// `JSON.stringify(x).length` counts UTF-16 code units, which undercounts
// any non-ASCII text (#223).

/** UTF-8 byte length of a string, without allocating an encoded copy. */
export function utf8ByteLength(s: string): number {
  let n = 0;
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c < 0x80) n += 1;
    else if (c < 0x800) n += 2;
    else if (c >= 0xd800 && c <= 0xdbff && i + 1 < s.length) {
      const d = s.charCodeAt(i + 1);
      // A surrogate pair is one 4-byte character; a lone high surrogate
      // encodes as U+FFFD (3 bytes).
      if (d >= 0xdc00 && d <= 0xdfff) {
        n += 4;
        i++;
      } else n += 3;
    } else n += 3;
  }
  return n;
}

/** Size of a parsed body as UTF-8 JSON; 0 for no body. */
export function jsonBytes(value: unknown): number {
  if (value === undefined || value === null) return 0;
  return utf8ByteLength(JSON.stringify(value));
}
