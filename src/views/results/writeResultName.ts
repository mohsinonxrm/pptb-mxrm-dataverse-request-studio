// The record name a write result card shows (#209).

/**
 * The record's primary name: from the response first (a computed primary
 * name such as a contact's `fullname` is never in the request body), then
 * from the value that was sent. Null when neither has one.
 */
export function writeResultName(
  primaryName: string | undefined,
  responseBody: unknown,
  sentValues: Record<string, unknown>,
): string | null {
  if (!primaryName) return null;
  const returned =
    responseBody && typeof responseBody === 'object'
      ? (responseBody as Record<string, unknown>)[primaryName]
      : undefined;
  if (typeof returned === 'string' && returned.trim()) return returned;
  const sent = sentValues[primaryName];
  return typeof sent === 'string' && sent.trim() ? sent : null;
}
