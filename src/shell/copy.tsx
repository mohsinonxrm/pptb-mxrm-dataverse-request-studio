// Copy buttons (#242). Every copy goes through ToolBox's clipboard bridge
// when embedded (`copyToClipboardSafe`): the webview can refuse
// `navigator.clipboard`, and a refused copy used to fail silently, leaving
// stale clipboard content to be pasted into Postman or a script. Every copy
// now says whether it worked.

import { copyToClipboardSafe } from '../host/pptbBridge';
import { notify } from './dialogs';

export interface CopyOptions {
  /**
   * Show a "Copied" toast on success. Turn off where the button itself shows
   * "Copied". A failure is always reported.
   */
  toastOnSuccess?: boolean;
  /**
   * The copied text is on screen, so a failure can suggest selecting it and
   * pressing Ctrl+C. Off for text that isn't shown (results as TSV, …).
   */
  shown?: boolean;
}

/**
 * Copy `text` and tell the user how it went. `what` names it inside a
 * sentence ("the URL", "3 rows"). Resolves true when copied; never rejects.
 */
export async function copyText(
  text: string,
  what: string,
  { toastOnSuccess = true, shown = true }: CopyOptions = {},
): Promise<boolean> {
  try {
    await copyToClipboardSafe(text);
  } catch (e) {
    const reason = e instanceof Error ? e.message.trim().replace(/\.$/, '') : '';
    const why = reason ? ` (${reason})` : '';
    const instead = shown ? 'Select it and press Ctrl+C instead.' : 'Try again in a moment.';
    notify(`Couldn't copy ${what}${why}. ${instead}`, {
      title: "Couldn't copy",
      intent: 'error',
    });
    return false;
  }
  if (toastOnSuccess) {
    notify(`Copied ${what} to the clipboard.`, {
      title: 'Copied',
      intent: 'success',
      timeout: 2500,
    });
  }
  return true;
}
