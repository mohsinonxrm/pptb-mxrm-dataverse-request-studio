// BinarySourceCard — payload picker for binary uploads.
//
// Source toggle: Upload file / From URL / Paste base64.
//
// We use a segmented Fluent v9 button group for the source toggle. When
// "Upload file" is active we keep only the file's name, size and type: the
// bytes are never read (#222). Execute can't send binary content from
// inside ToolBox and the Code tab uses placeholders, so a base64 copy in
// state was only a way for a large file to freeze the tool. A mode that
// previews the file gets the File itself through `onFile`. When "Paste
// base64" we let the user paste raw base64, up to MAX_PASTED_BASE64_BYTES.
// When "From URL" we just store the URL (informational — actual fetch would
// happen client-side at execute time).

import { useEffect, useRef, useState } from 'react';
import {
  Field,
  Input,
  Button,
  Textarea,
  Caption1,
  Badge,
  tokens,
  mergeClasses,
  Tooltip,
} from '@fluentui/react-components';
import {
  ArrowUpload20Regular,
  Link20Regular,
  Code20Regular,
  Document20Filled,
  DocumentArrowUp20Filled,
} from '@fluentui/react-icons';
import { useStudioStyles } from '../primitives/styles';
import { PaneHead } from './PaneHead';
import { formatSize } from '../engine/binaryBuilders';
import type { BinarySource } from '../state/binaryState';
import type { RequestGroup } from '../registry/requestTypes';

export interface BinarySourceCardProps {
  source: BinarySource;
  setSource: (s: BinarySource) => void;
  fileName: string;
  setFileName: (n: string) => void;
  fileSize: number;
  setFileSize: (n: number) => void;
  mimeType: string;
  setMimeType: (m: string) => void;
  bodyBase64: string;
  setBodyBase64: (b: string) => void;
  bodyUrl: string;
  setBodyUrl: (u: string) => void;
  /** Restrict file picker to these MIME types (used for image uploads). */
  accept?: string;
  group?: RequestGroup;
  /** The chosen file, for a preview (object URL); null when cleared. */
  onFile?: (file: File | null) => void;
}

/** Largest pasted base64 body kept, decoded size (#222). */
export const MAX_PASTED_BASE64_BYTES = 10 * 1024 * 1024;

export function BinarySourceCard(props: BinarySourceCardProps) {
  const s = useStudioStyles();
  const inputRef = useRef<HTMLInputElement>(null);
  const [pasteError, setPasteError] = useState<string | null>(null);
  // A refused paste's message belongs to that paste, not to the next source.
  useEffect(() => setPasteError(null), [props.source.kind]);

  return (
    <div className={mergeClasses(s.inlineCard)} style={{ padding: 12, maxWidth: 880 }}>
      <PaneHead
        icon={DocumentArrowUp20Filled}
        title="Payload source"
        sub="Where the bytes come from. Switch source to change how the body gets composed."
        group={props.group ?? 'binary'}
      />

      {/* Segmented source toggle */}
      <div
        style={{
          display: 'inline-flex',
          border: `1px solid ${tokens.colorNeutralStroke2}`,
          borderRadius: tokens.borderRadiusSmall,
          overflow: 'hidden',
          marginBottom: 14,
        }}
      >
        <SourceTab
          active={props.source.kind === 'file'}
          onClick={() => props.setSource({ kind: 'file' })}
          icon={<ArrowUpload20Regular />}
          label="Upload file"
        />
        <SourceTab
          active={props.source.kind === 'url'}
          onClick={() => props.setSource({ kind: 'url' })}
          icon={<Link20Regular />}
          label="From URL"
        />
        <SourceTab
          active={props.source.kind === 'base64'}
          onClick={() => props.setSource({ kind: 'base64' })}
          icon={<Code20Regular />}
          label="Paste base64"
        />
      </div>

      {props.source.kind === 'file' && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 10, maxWidth: 560 }}>
          <Field label="File">
            <input
              ref={inputRef}
              type="file"
              accept={props.accept}
              style={{ fontSize: 12 }}
              onChange={(e) => {
                const f = e.currentTarget.files?.[0];
                if (!f) return;
                // Name, size and type only — never the bytes (#222).
                props.setFileName(f.name);
                props.setFileSize(f.size);
                if (f.type) props.setMimeType(f.type);
                props.setBodyBase64('');
                props.onFile?.(f);
              }}
            />
          </Field>
          {props.fileName && (
            <div
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: 8,
                padding: '8px 10px',
                background: tokens.colorNeutralBackground2,
                borderRadius: tokens.borderRadiusSmall,
                fontSize: 11,
              }}
            >
              <Document20Filled
                style={{ color: tokens.colorBrandForeground1, width: 18, height: 18 }}
              />
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontWeight: 600 }}>{props.fileName}</div>
                <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                  {formatSize(props.fileSize)} · {props.mimeType || 'application/octet-stream'}
                </Caption1>
              </div>
              <Tooltip content="Pick a different file" relationship="label">
                <Button size="small" appearance="subtle" onClick={() => inputRef.current?.click()}>
                  Change
                </Button>
              </Tooltip>
            </div>
          )}
        </div>
      )}

      {props.source.kind === 'url' && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 10, maxWidth: 720 }}>
          <Field
            label="URL"
            hint={
              <span style={{ color: tokens.colorNeutralForeground3 }}>
                Informational — the studio doesn't fetch external URLs at execute time. Use this
                when the file lives on a CDN and you'll fetch it client-side before calling
                Dataverse.
              </span>
            }
          >
            <Input
              value={props.bodyUrl}
              onChange={(_, d) => props.setBodyUrl(d.value)}
              placeholder="https://cdn.example.com/file.pdf"
              style={{ fontFamily: tokens.fontFamilyMonospace }}
            />
          </Field>
          <Field label="File name">
            <Input
              value={props.fileName}
              onChange={(_, d) => props.setFileName(d.value)}
              placeholder="file.pdf"
            />
          </Field>
          <Field label="MIME type">
            <Input
              value={props.mimeType}
              onChange={(_, d) => props.setMimeType(d.value)}
              placeholder="application/pdf"
              style={{ fontFamily: tokens.fontFamilyMonospace }}
            />
          </Field>
        </div>
      )}

      {props.source.kind === 'base64' && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
          <Field label="File name">
            <Input
              value={props.fileName}
              onChange={(_, d) => props.setFileName(d.value)}
              placeholder="file.bin"
            />
          </Field>
          <Field label="MIME type">
            <Input
              value={props.mimeType}
              onChange={(_, d) => props.setMimeType(d.value)}
              placeholder="application/octet-stream"
              style={{ fontFamily: tokens.fontFamilyMonospace }}
            />
          </Field>
          <Field
            label="Base64 body"
            validationState={pasteError ? 'error' : 'none'}
            validationMessage={pasteError ?? undefined}
            hint={
              <span>
                Raw base64-encoded bytes.{' '}
                <Badge appearance="ghost">{formatSize(estimateBase64Size(props.bodyBase64))}</Badge>{' '}
                after decode.
              </span>
            }
          >
            <Textarea
              rows={6}
              value={props.bodyBase64}
              onChange={(_, d) => {
                // A huge paste would freeze the tool; the Code tab uses a
                // placeholder for the body anyway (#222).
                const size = estimateBase64Size(d.value);
                if (size > MAX_PASTED_BASE64_BYTES) {
                  setPasteError(
                    `That's ${formatSize(size)} decoded; up to ${formatSize(MAX_PASTED_BASE64_BYTES)} can be pasted. Use Upload file for larger files.`,
                  );
                  return;
                }
                setPasteError(null);
                props.setBodyBase64(d.value);
                props.setFileSize(size);
              }}
              style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 11 }}
              placeholder="JVBERi0xLjQK…"
            />
          </Field>
        </div>
      )}
    </div>
  );
}

function SourceTab({
  active,
  onClick,
  icon,
  label,
}: {
  active: boolean;
  onClick: () => void;
  icon: React.ReactElement;
  label: string;
}) {
  return (
    <Button
      size="small"
      appearance={active ? 'primary' : 'subtle'}
      icon={icon}
      onClick={onClick}
      style={{ borderRadius: 0, border: 'none' }}
    >
      {label}
    </Button>
  );
}

/**
 * Decoded size of pasted base64. Line breaks and other whitespace (base64 is
 * often wrapped at 76 characters) and a `data:…;base64,` prefix aren't data.
 */
export function estimateBase64Size(text: string): number {
  const b64 = text.replace(/^data:[^,]*,/, '').replace(/\s+/g, '');
  if (!b64) return 0;
  const padding = (b64.match(/=+$/) ?? [''])[0].length;
  return Math.max(0, Math.floor((b64.length * 3) / 4) - padding);
}
