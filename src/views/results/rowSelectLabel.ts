// The accessible name of a results row's selection checkbox (#387): the
// record's primary name when the row has one, else its position.

export function rowSelectLabel(
  item: Record<string, unknown>,
  primaryName: string | undefined,
  index: number,
): string {
  const name = primaryName ? item[primaryName] : undefined;
  return typeof name === 'string' && name.trim()
    ? `Select ${name.trim()}`
    : `Select row ${index + 1}`;
}
