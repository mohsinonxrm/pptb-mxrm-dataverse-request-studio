// BinarySourceCard — payload picker for binary uploads.
//
// Source toggle: Upload file / From URL / Paste base64.
//
// We use a segmented Fluent v9 button group for the source toggle. When
// "Upload file" is active we expose a file input + read it into memory as
// base64 + bytes-count metadata. When "Paste base64" we let the user paste
// raw base64. When "From URL" we just store the URL (informational — actual
// fetch would happen client-side at execute time).

import { useRef } from 'react';
import {
  Field, Input, Button, Textarea, Caption1, Badge, tokens, mergeClasses, Tooltip,
} from '@fluentui/react-components';
import { ArrowUpload20Regular, Link20Regular, Code20Regular, Document20Filled, DocumentArrowUp20Filled } from '@fluentui/react-icons';
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
}

export function BinarySourceCard(props: BinarySourceCardProps) {
  const s = useStudioStyles();
  const inputRef = useRef<HTMLInputElement>(null);

  return (
    <div className={mergeClasses(s.inlineCard)} style={{ padding: 12, maxWidth: 880 }}>
      <PaneHead
        icon={DocumentArrowUp20Filled}
        title="Payload source"
        sub="Where the bytes come from. Switch source to change how the body gets composed."
        group={props.group ?? 'binary'}
      />

      {/* Segmented source toggle */}
      <div style={{ display: 'inline-flex', border: `1px solid ${tokens.colorNeutralStroke2}`, borderRadius: tokens.borderRadiusSmall, overflow: 'hidden', marginBottom: 14 }}>
        <SourceTab active={props.source.kind === 'file'}   onClick={() => props.setSource({ kind: 'file' })}   icon={<ArrowUpload20Regular />} label="Upload file" />
        <SourceTab active={props.source.kind === 'url'}    onClick={() => props.setSource({ kind: 'url' })}    icon={<Link20Regular />}        label="From URL" />
        <SourceTab active={props.source.kind === 'base64'} onClick={() => props.setSource({ kind: 'base64' })} icon={<Code20Regular />}        label="Paste base64" />
      </div>

      {props.source.kind === 'file' && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 10, maxWidth: 560 }}>
          <Field label="File">
            <input
              ref={inputRef}
              type="file"
              accept={props.accept}
              style={{ fontSize: 12 }}
              onChange={async (e) => {
                const f = e.currentTarget.files?.[0];
                if (!f) return;
                props.setFileName(f.name);
                props.setFileSize(f.size);
                if (f.type) props.setMimeType(f.type);
                // Read base64 — strip the data: prefix so the wire payload is pure base64
                const reader = new FileReader();
                reader.onload = () => {
                  const result = String(reader.result ?? '');
                  const b64 = result.startsWith('data:') ? result.split(',')[1] ?? '' : result;
                  props.setBodyBase64(b64);
                };
                reader.readAsDataURL(f);
              }}
            />
          </Field>
          {props.fileName && (
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '8px 10px', background: tokens.colorNeutralBackground2, borderRadius: tokens.borderRadiusSmall, fontSize: 11 }}>
              <Document20Filled style={{ color: tokens.colorBrandForeground1, width: 18, height: 18 }} />
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ fontWeight: 600 }}>{props.fileName}</div>
                <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                  {formatSize(props.fileSize)} · {props.mimeType || 'application/octet-stream'}
                </Caption1>
              </div>
              <Tooltip content="Pick a different file" relationship="label">
                <Button size="small" appearance="subtle" onClick={() => inputRef.current?.click()}>Change</Button>
              </Tooltip>
            </div>
          )}
        </div>
      )}

      {props.source.kind === 'url' && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 10, maxWidth: 720 }}>
          <Field
            label="URL"
            hint={<span style={{ color: tokens.colorNeutralForeground3 }}>Informational — the studio doesn't fetch external URLs at execute time. Use this when the file lives on a CDN and you'll fetch it client-side before calling Dataverse.</span>}
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
            <Input value={props.fileName} onChange={(_, d) => props.setFileName(d.value)} placeholder="file.bin" />
          </Field>
          <Field label="MIME type">
            <Input value={props.mimeType} onChange={(_, d) => props.setMimeType(d.value)} placeholder="application/octet-stream" style={{ fontFamily: tokens.fontFamilyMonospace }} />
          </Field>
          <Field
            label="Base64 body"
            hint={<span>Raw base64-encoded bytes. <Badge appearance="ghost">{formatSize(estimateBase64Size(props.bodyBase64))}</Badge> after decode.</span>}
          >
            <Textarea
              rows={6}
              value={props.bodyBase64}
              onChange={(_, d) => {
                props.setBodyBase64(d.value);
                props.setFileSize(estimateBase64Size(d.value));
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

function SourceTab({ active, onClick, icon, label }: { active: boolean; onClick: () => void; icon: React.ReactElement; label: string }) {
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

function estimateBase64Size(b64: string): number {
  if (!b64) return 0;
  const padding = (b64.match(/=+$/) ?? [''])[0].length;
  return Math.max(0, Math.floor((b64.length * 3) / 4) - padding);
}
