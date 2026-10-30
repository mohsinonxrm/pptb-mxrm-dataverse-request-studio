// The clipboard bridge (#242): ToolBox's copy first when embedded, the
// browser's otherwise, and an error when nothing could copy (it used to
// resolve as if it had).
import { afterEach, describe, expect, it, vi } from 'vitest';

const win = window as unknown as { toolboxAPI?: unknown };
const nav = navigator as unknown as { clipboard?: unknown };
const realClipboard = Object.getOwnPropertyDescriptor(navigator, 'clipboard');

function setBrowserClipboard(clipboard: unknown) {
  Object.defineProperty(navigator, 'clipboard', { value: clipboard, configurable: true });
}

/** Load the bridge afresh: it decides once, at load, whether it's embedded. */
async function loadBridge(hostCopy?: (t: string) => Promise<void>) {
  vi.resetModules();
  if (hostCopy) {
    win.toolboxAPI = {
      utils: { copyToClipboard: hostCopy, getCurrentTheme: async () => 'light' },
      connections: { getActiveConnection: async () => null },
      events: { on: () => {} },
    };
  }
  return import('./pptbBridge');
}

afterEach(() => {
  delete win.toolboxAPI;
  if (realClipboard) Object.defineProperty(navigator, 'clipboard', realClipboard);
  else delete nav.clipboard;
  vi.restoreAllMocks();
});

describe('copyToClipboardSafe', () => {
  it("uses ToolBox's clipboard when embedded", async () => {
    const hostCopy = vi.fn(async () => {});
    const writeText = vi.fn(async () => {});
    setBrowserClipboard({ writeText });
    const { copyToClipboardSafe } = await loadBridge(hostCopy);
    await copyToClipboardSafe('https://x');
    expect(hostCopy).toHaveBeenCalledWith('https://x');
    expect(writeText).not.toHaveBeenCalled();
  });

  it("falls back to the browser's when ToolBox's fails", async () => {
    vi.spyOn(console, 'warn').mockImplementation(() => {});
    const writeText = vi.fn(async () => {});
    setBrowserClipboard({ writeText });
    const { copyToClipboardSafe } = await loadBridge(async () => {
      throw new Error('IPC refused');
    });
    await copyToClipboardSafe('code');
    expect(writeText).toHaveBeenCalledWith('code');
  });

  it("rejects with ToolBox's error when there's no browser clipboard to fall back to", async () => {
    vi.spyOn(console, 'warn').mockImplementation(() => {});
    setBrowserClipboard(undefined);
    const { copyToClipboardSafe } = await loadBridge(async () => {
      throw new Error('IPC refused');
    });
    await expect(copyToClipboardSafe('code')).rejects.toThrow('IPC refused');
  });

  it('rejects when the browser refuses, standalone', async () => {
    setBrowserClipboard({ writeText: vi.fn(async () => Promise.reject(new Error('denied'))) });
    const { copyToClipboardSafe } = await loadBridge();
    await expect(copyToClipboardSafe('x')).rejects.toThrow('denied');
  });

  it('rejects when there is no clipboard at all, instead of resolving', async () => {
    setBrowserClipboard(undefined);
    const { copyToClipboardSafe } = await loadBridge();
    await expect(copyToClipboardSafe('x')).rejects.toThrow('no clipboard is available');
  });
});

describe('readClipboardText', () => {
  it('returns the text', async () => {
    setBrowserClipboard({ readText: vi.fn(async () => 'https://org/api/data/v9.2/accounts') });
    const { readClipboardText } = await loadBridge();
    expect(await readClipboardText()).toBe('https://org/api/data/v9.2/accounts');
  });

  it('returns null when the read is refused or unavailable', async () => {
    setBrowserClipboard({ readText: vi.fn(async () => Promise.reject(new Error('focus'))) });
    let bridge = await loadBridge();
    expect(await bridge.readClipboardText()).toBeNull();
    setBrowserClipboard(undefined);
    bridge = await loadBridge();
    expect(await bridge.readClipboardText()).toBeNull();
  });
});
