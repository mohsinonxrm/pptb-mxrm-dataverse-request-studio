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
    expect(await copyText('https://x', 'the URL')).toBe(true);
    expect(bridge.copy).toHaveBeenCalledWith('https://x');
    expect(dialogs.notify).toHaveBeenCalledWith('Copied the URL to the clipboard.', {
      title: 'Copied',
      intent: 'success',
      timeout: 2500,
    });
  });

  it('reads right for a plural', async () => {
    bridge.copy.mockResolvedValue(undefined);
    await copyText('a\tb', '3 rows', { shown: false });
    expect(dialogs.notify.mock.calls[0][0]).toBe('Copied 3 rows to the clipboard.');
  });

  it('says so when the copy fails, and never rejects', async () => {
    bridge.copy.mockRejectedValue(new Error('no clipboard is available'));
    expect(await copyText('code', 'the code')).toBe(false);
    expect(dialogs.notify).toHaveBeenCalledWith(
      "Couldn't copy the code (no clipboard is available). Select it and press Ctrl+C instead.",
      { title: "Couldn't copy", intent: 'error' },
    );
  });

  it("doesn't double the reason's own full stop", async () => {
    bridge.copy.mockRejectedValue(new Error('Write permission denied.'));
    await copyText('x', 'the URL');
    expect(dialogs.notify.mock.calls[0][0]).toBe(
      "Couldn't copy the URL (Write permission denied). Select it and press Ctrl+C instead.",
    );
  });

  it('leaves out a reason it has no message for', async () => {
    bridge.copy.mockRejectedValue('denied');
    await copyText('x', 'the job ID');
    expect(dialogs.notify.mock.calls[0][0]).toBe(
      "Couldn't copy the job ID. Select it and press Ctrl+C instead.",
    );
  });

  it("doesn't suggest selecting text that isn't on screen", async () => {
    bridge.copy.mockRejectedValue(new Error('denied'));
    await copyText('a\tb', '3 rows', { shown: false });
    expect(dialogs.notify.mock.calls[0][0]).toBe(
      "Couldn't copy 3 rows (denied). Try again in a moment.",
    );
  });

  it('skips the success toast where the button shows "Copied", but still reports a failure', async () => {
    bridge.copy.mockResolvedValue(undefined);
    expect(await copyText('ids', 'the IDs', { toastOnSuccess: false })).toBe(true);
    expect(dialogs.notify).not.toHaveBeenCalled();
    bridge.copy.mockRejectedValue(new Error('denied'));
    expect(await copyText('ids', 'the IDs', { toastOnSuccess: false })).toBe(false);
    expect(dialogs.notify).toHaveBeenCalledTimes(1);
  });
});
