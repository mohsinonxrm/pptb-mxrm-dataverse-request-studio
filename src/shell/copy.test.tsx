// Copy buttons say whether the copy worked (#242).
import { beforeEach, describe, expect, it, vi } from 'vitest';

const bridge = vi.hoisted(() => ({ copy: vi.fn() }));
vi.mock('../host/pptbBridge', () => ({ copyToClipboardSafe: bridge.copy }));
const dialogs = vi.hoisted(() => ({ notify: vi.fn() }));
vi.mock('./dialogs', () => ({ notify: dialogs.notify }));

import { copyText } from './copy';

beforeEach(() => {
  bridge.copy.mockReset();
  dialogs.notify.mockReset();
});

describe('copyText', () => {
  it('copies and shows a short "Copied" toast', async () => {
    bridge.copy.mockResolvedValue(undefined);
    expect(await copyText('https://x', 'URL')).toBe(true);
    expect(bridge.copy).toHaveBeenCalledWith('https://x');
    expect(dialogs.notify).toHaveBeenCalledWith('URL is on the clipboard.', {
      title: 'Copied',
      intent: 'success',
      timeout: 2500,
    });
  });

  it('says so when the copy fails, and never rejects', async () => {
    bridge.copy.mockRejectedValue(new Error('no clipboard is available'));
    expect(await copyText('code', 'Code')).toBe(false);
    expect(dialogs.notify).toHaveBeenCalledWith(
      "Code wasn't copied (no clipboard is available). Select it and press Ctrl+C instead.",
      { title: "Couldn't copy", intent: 'error' },
    );
  });

  it("doesn't double the reason's own full stop", async () => {
    bridge.copy.mockRejectedValue(new Error('Write permission denied.'));
    await copyText('x', 'URL');
    expect(dialogs.notify.mock.calls[0][0]).toBe(
      "URL wasn't copied (Write permission denied). Select it and press Ctrl+C instead.",
    );
  });

  it('leaves out a reason it has no message for', async () => {
    bridge.copy.mockRejectedValue('denied');
    await copyText('x', 'Job id');
    expect(dialogs.notify.mock.calls[0][0]).toBe(
      "Job id wasn't copied. Select it and press Ctrl+C instead.",
    );
  });

  it('skips the success toast where the button shows "Copied", but still reports a failure', async () => {
    bridge.copy.mockResolvedValue(undefined);
    expect(await copyText('ids', 'Ids', { toastOnSuccess: false })).toBe(true);
    expect(dialogs.notify).not.toHaveBeenCalled();
    bridge.copy.mockRejectedValue(new Error('denied'));
    expect(await copyText('ids', 'Ids', { toastOnSuccess: false })).toBe(false);
    expect(dialogs.notify).toHaveBeenCalledTimes(1);
  });
});
