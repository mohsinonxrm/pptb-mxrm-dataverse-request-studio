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
}

/**
 * Copy `text` and tell the user how it went. `what` names it in the
 * messages ("URL", "Code", …). Resolves true when copied; never rejects.
 */
export async function copyText(
  text: string,
  what: string,
  { toastOnSuccess = true }: CopyOptions = {},
): Promise<boolean> {
  try {
    await copyToClipboardSafe(text);
  } catch (e) {
    const reason = e instanceof Error ? e.message.trim().replace(/\.$/, '') : '';
    const why = reason ? ` (${reason})` : '';
    notify(`${what} wasn't copied${why}. Select it and press Ctrl+C instead.`, {
      title: "Couldn't copy",
      intent: 'error',
    });
    return false;
  }
  if (toastOnSuccess) {
    notify(`${what} is on the clipboard.`, { title: 'Copied', intent: 'success', timeout: 2500 });
  }
  return true;
}
