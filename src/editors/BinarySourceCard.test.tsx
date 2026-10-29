// #222: choosing a file keeps its name, size and type, never its bytes; a
// pasted base64 body is capped.
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { BinarySourceCard, MAX_PASTED_BASE64_BYTES } from './BinarySourceCard';
import type { BinarySource } from '../state/binaryState';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let root: Root | null = null;
let host: HTMLDivElement | null = null;

afterEach(() => {
  act(() => root?.unmount());
  host?.remove();
  root = null;
  host = null;
  vi.restoreAllMocks();
});

function mount(source: BinarySource) {
  const calls = {
    fileName: vi.fn(),
    fileSize: vi.fn(),
    mimeType: vi.fn(),
    bodyBase64: vi.fn(),
    onFile: vi.fn(),
  };
  host = document.createElement('div');
  document.body.appendChild(host);
  root = createRoot(host);
  act(() =>
    root!.render(
      <FluentProvider theme={webLightTheme}>
        <BinarySourceCard
          source={source}
          setSource={() => undefined}
          fileName=""
          setFileName={calls.fileName}
          fileSize={0}
          setFileSize={calls.fileSize}
          mimeType=""
          setMimeType={calls.mimeType}
          bodyBase64=""
          setBodyBase64={calls.bodyBase64}
          bodyUrl=""
          setBodyUrl={() => undefined}
          onFile={calls.onFile}
        />
      </FluentProvider>,
    ),
  );
  return calls;
}

describe('BinarySourceCard', () => {
  it('a chosen file is described, not read', () => {
    const read = vi.spyOn(FileReader.prototype, 'readAsDataURL');
    const calls = mount({ kind: 'file' });
    const file = new File([new Uint8Array(2048)], '报告.pdf', { type: 'application/pdf' });
    const input = host!.querySelector('input[type=file]') as HTMLInputElement;
    Object.defineProperty(input, 'files', { value: [file] });
    act(() => {
      input.dispatchEvent(new Event('change', { bubbles: true }));
    });
    expect(calls.fileName).toHaveBeenCalledWith('报告.pdf');
    expect(calls.fileSize).toHaveBeenCalledWith(2048);
    expect(calls.mimeType).toHaveBeenCalledWith('application/pdf');
    expect(calls.bodyBase64).toHaveBeenCalledWith('');
    expect(calls.onFile).toHaveBeenCalledWith(file);
    expect(read).not.toHaveBeenCalled();
  });

  it('refuses a pasted body over the cap, with a message', () => {
    const calls = mount({ kind: 'base64' });
    const area = host!.querySelector('textarea') as HTMLTextAreaElement;
    const paste = (value: string) =>
      act(() => {
        Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!.call(
          area,
          value,
        );
        area.dispatchEvent(new Event('input', { bubbles: true }));
      });
    paste('QUJD');
    expect(calls.bodyBase64).toHaveBeenLastCalledWith('QUJD');
    expect(calls.fileSize).toHaveBeenLastCalledWith(3);
    calls.bodyBase64.mockClear();
    paste('A'.repeat(Math.ceil((MAX_PASTED_BASE64_BYTES * 4) / 3) + 8));
    expect(calls.bodyBase64).not.toHaveBeenCalled();
    expect(host!.textContent).toMatch(/can be pasted\. Use Upload file for larger files\./);
  });
});

describe('estimateBase64Size', () => {
  it('ignores line breaks, other whitespace and a data: prefix', async () => {
    const { estimateBase64Size } = await import('./BinarySourceCard');
    expect(estimateBase64Size('QUJD')).toBe(3);
    expect(estimateBase64Size('QUJD\r\nREVG\r\n')).toBe(6);
    expect(estimateBase64Size('QQ==\n')).toBe(1);
    expect(estimateBase64Size('data:image/png;base64,QUJD')).toBe(3);
  });

  it('a wrapped body just under the cap counts as under it', async () => {
    const { estimateBase64Size } = await import('./BinarySourceCard');
    const chars = Math.floor((MAX_PASTED_BASE64_BYTES * 4) / 3 / 4) * 4;
    const wrapped = ('A'.repeat(chars).match(/.{1,76}/g) ?? []).join('\r\n');
    expect(estimateBase64Size(wrapped)).toBeLessThanOrEqual(MAX_PASTED_BASE64_BYTES);
  });
});
