// Open a returned record in the model-driven app (#293).
//
// Inside ToolBox the URL goes through `toolboxAPI.utils.openInConnectionBrowser`,
// which opens it in the browser profile set on the active connection, so the
// user is already signed in. Older ToolBox builds without it, and standalone,
// fall back to a new browser window.

const GUID = /^\{?([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})\}?$/i;

/**
 * The record's form URL, or null when there's no environment, table or
 * record GUID (aggregates and `$apply` rows have no primary ID).
 */
export function recordFormUrl(
  host: string | undefined,
  table: string | undefined,
  id: unknown,
): string | null {
  const guid = typeof id === 'string' ? GUID.exec(id.trim())?.[1] : undefined;
  const etn = table?.trim().toLowerCase();
  if (!host || !etn || !guid) return null;
  return `https://${host}/main.aspx?pagetype=entityrecord&etn=${encodeURIComponent(etn)}&id=${guid.toLowerCase()}`;
}

/** Open `url` in the connection's browser, or a new window outside ToolBox. */
export async function openInBrowser(url: string): Promise<void> {
  const open =
    typeof window !== 'undefined' ? window.toolboxAPI?.utils?.openInConnectionBrowser : undefined;
  if (typeof open === 'function') {
    try {
      await open(url);
      return;
    } catch (e) {
      console.warn('[openInDataverse] openInConnectionBrowser failed, opening a window instead', e);
    }
  }
  window.open(url, '_blank', 'noopener');
}
