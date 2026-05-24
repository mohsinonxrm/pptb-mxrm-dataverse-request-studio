// FileColumnPicker / ImageColumnPicker — pick a File-typed or Image-typed
// column on the target table. The picker surfaces the column's max size,
// mime / image constraints, and (for image columns) the primary vs custom
// + canStoreFullImage distinction.

import { useMemo } from 'react';
import { Field, Combobox, Option, Caption1, Badge, tokens, MessageBar, MessageBarBody } from '@fluentui/react-components';
import { Document20Filled, Image20Filled } from '@fluentui/react-icons';
import { PaneHead } from './PaneHead';
import { findTable, type FileColumnMeta, type ImageColumnMeta } from '../mock/metadata';
import { formatSize } from '../engine/binaryBuilders';
import { fileNameColumnFor, imageCompanionsFor } from '../state/binaryState';
import type { RequestGroup } from '../registry/requestTypes';

// ──────────────────────────────────────────────────────────────
// File column picker
// ──────────────────────────────────────────────────────────────
export interface FileColumnPickerProps {
  table: string;
  value: string | null;
  onChange: (col: string | null) => void;
  group?: RequestGroup;
}

export function FileColumnPicker({ table, value, onChange, group = 'binary' }: FileColumnPickerProps) {
  const tbl = findTable(table);
  const fileCols = useMemo(
    () => (tbl?.columns ?? []).filter((c): c is FileColumnMeta => c.attributeType === 'File'),
    [tbl],
  );
  const cur = fileCols.find(c => c.logicalName === value);

  if (!tbl) {
    return (
      <MessageBar layout="multiline" intent="error">
        <MessageBarBody>Unknown table <code>{table}</code>.</MessageBarBody>
      </MessageBar>
    );
  }
  if (fileCols.length === 0) {
    return (
      <div>
        <PaneHead icon={Document20Filled} title="File column" group={group} />
        <MessageBar layout="multiline" intent="warning" style={{ maxWidth: 720 }}>
          <MessageBarBody>
            <strong>{tbl.displayName}</strong> has no File-typed columns. Add one via the maker portal or pick a different table.
          </MessageBarBody>
        </MessageBar>
      </div>
    );
  }

  return (
    <div>
      <PaneHead
        icon={Document20Filled}
        title="File column"
        sub="Pick the File-typed column to operate on. Max size + target are sourced from the column metadata."
        group={group}
      />
      <div style={{ maxWidth: 720 }}>
        <Field label="Column">
          <Combobox
            value={cur?.displayName ?? cur?.logicalName ?? ''}
            selectedOptions={value ? [value] : []}
            onOptionSelect={(_, d) => onChange(d.optionValue ?? null)}
            placeholder="Pick a file column…"
          >
            {fileCols.map(c => (
              <Option key={c.logicalName} value={c.logicalName} text={c.displayName}>
                <FileColumnOption col={c} />
              </Option>
            ))}
          </Combobox>
        </Field>

        {cur && (
          <div style={{ marginTop: 16 }}>
            <FileColumnSummary col={cur} />
          </div>
        )}
      </div>
    </div>
  );
}

function FileColumnOption({ col }: { col: FileColumnMeta }) {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', minWidth: 0 }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
        <span style={{ fontFamily: tokens.fontFamilyMonospace, fontWeight: 600 }}>{col.logicalName}</span>
        <Badge appearance="ghost" size="extra-small">File</Badge>
        <Badge appearance="ghost" size="extra-small" style={{ marginLeft: 'auto' }}>
          max {formatSize((col.maxSizeInKB ?? 32768) * 1024)}
        </Badge>
      </div>
      <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
        {col.displayName}
        {col.binaryTarget && col.binaryTarget !== 'file' && <> · {col.binaryTarget} body</>}
      </Caption1>
    </div>
  );
}

function FileColumnSummary({ col }: { col: FileColumnMeta }) {
  const maxBytes = (col.maxSizeInKB ?? 32768) * 1024;
  return (
    <div style={{
      border: `1px solid ${tokens.colorNeutralStroke2}`,
      borderRadius: tokens.borderRadiusMedium,
      padding: 12,
      background: tokens.colorNeutralBackground1,
    }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
        <Document20Filled style={{ color: tokens.colorBrandForeground1 }} />
        <strong style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 13 }}>{col.logicalName}</strong>
        <Badge appearance="tint" color="brand">File</Badge>
        <span style={{ flexGrow: 1 }} />
        <Caption1 style={{ color: tokens.colorNeutralForeground3, fontFamily: tokens.fontFamilyMonospace }}>
          {col.displayName}
        </Caption1>
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: '160px 1fr', rowGap: 4, fontSize: 11 }}>
        <span style={{ color: tokens.colorNeutralForeground3 }}>Max file size</span>
        <span>{formatSize(maxBytes)} <Caption1 style={{ color: tokens.colorNeutralForeground3, marginLeft: 4 }}>· {(col.maxSizeInKB ?? 32768).toLocaleString()} KB</Caption1></span>
        <span style={{ color: tokens.colorNeutralForeground3 }}>Companion column</span>
        <span style={{ fontFamily: tokens.fontFamilyMonospace }}>{fileNameColumnFor(col.logicalName)} <Caption1 style={{ color: tokens.colorNeutralForeground3, marginLeft: 4, fontFamily: tokens.fontFamilyBase }}>· read-only string — $select to get the filename without downloading bytes</Caption1></span>
        <span style={{ color: tokens.colorNeutralForeground3 }}>Retrieve returns</span>
        <span>file id (Guid). Pass to DeleteFile or compose download URL.</span>
        <span style={{ color: tokens.colorNeutralForeground3 }}>Binary target</span>
        <span><Badge appearance="ghost">{col.binaryTarget ?? 'file'}</Badge></span>
      </div>
    </div>
  );
}

// ──────────────────────────────────────────────────────────────
// Image column picker
// ──────────────────────────────────────────────────────────────
export interface ImageColumnPickerProps {
  table: string;
  value: string | null;
  onChange: (col: string | null) => void;
  group?: RequestGroup;
}

export function ImageColumnPicker({ table, value, onChange, group = 'binary' }: ImageColumnPickerProps) {
  const tbl = findTable(table);
  const imgCols = useMemo(
    () => (tbl?.columns ?? []).filter((c): c is ImageColumnMeta => c.attributeType === 'Image'),
    [tbl],
  );
  const cur = imgCols.find(c => c.logicalName === value);

  if (!tbl) {
    return (
      <MessageBar layout="multiline" intent="error">
        <MessageBarBody>Unknown table <code>{table}</code>.</MessageBarBody>
      </MessageBar>
    );
  }
  if (imgCols.length === 0) {
    return (
      <div>
        <PaneHead icon={Image20Filled} title="Image column" group={group} />
        <MessageBar layout="multiline" intent="warning" style={{ maxWidth: 720 }}>
          <MessageBarBody>
            <strong>{tbl.displayName}</strong> has no Image-typed columns. Tables get one when their primary image is configured; custom image columns are created via solution explorer.
          </MessageBarBody>
        </MessageBar>
      </div>
    );
  }

  return (
    <div>
      <PaneHead
        icon={Image20Filled}
        title="Image column"
        sub="Pick the Image-typed column. Primary image (entityimage) is thumbnail-only; custom image columns can store full size when configured."
        group={group}
      />
      <div style={{ maxWidth: 720 }}>
        <Field label="Column">
          <Combobox
            value={cur?.displayName ?? cur?.logicalName ?? ''}
            selectedOptions={value ? [value] : []}
            onOptionSelect={(_, d) => onChange(d.optionValue ?? null)}
            placeholder="Pick an image column…"
          >
            {imgCols.map(c => (
              <Option key={c.logicalName} value={c.logicalName} text={c.displayName}>
                <ImageColumnOption col={c} />
              </Option>
            ))}
          </Combobox>
        </Field>

        {cur && (
          <div style={{ marginTop: 16 }}>
            <ImageColumnSummary col={cur} />
          </div>
        )}
      </div>
    </div>
  );
}

function ImageColumnOption({ col }: { col: ImageColumnMeta }) {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', minWidth: 0 }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
        <span style={{ fontFamily: tokens.fontFamilyMonospace, fontWeight: 600 }}>{col.logicalName}</span>
        <Badge appearance="ghost" size="extra-small">Image</Badge>
        {col.isPrimaryImage && <Badge appearance="tint" color="brand" size="extra-small">primary</Badge>}
        {col.canStoreFullImage && <Badge appearance="tint" color="success" size="extra-small">full-size</Badge>}
        <Badge appearance="ghost" size="extra-small" style={{ marginLeft: 'auto' }}>
          {col.isPrimaryImage ? '144×144' : `max ${formatSize((col.maxSizeInKB ?? 10240) * 1024)}`}
        </Badge>
      </div>
      <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>{col.displayName}</Caption1>
    </div>
  );
}

function ImageColumnSummary({ col }: { col: ImageColumnMeta }) {
  const maxBytes = (col.maxSizeInKB ?? 10240) * 1024;
  return (
    <div style={{
      border: `1px solid ${tokens.colorNeutralStroke2}`,
      borderRadius: tokens.borderRadiusMedium,
      padding: 12,
      background: tokens.colorNeutralBackground1,
    }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
        <Image20Filled style={{ color: tokens.colorBrandForeground1 }} />
        <strong style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 13 }}>{col.logicalName}</strong>
        <Badge appearance="tint" color="brand">Image</Badge>
        {col.isPrimaryImage && <Badge appearance="tint" color="brand">primary</Badge>}
        {col.canStoreFullImage
          ? <Badge appearance="tint" color="success">full-size enabled</Badge>
          : <Badge appearance="tint" color="subtle">thumbnail only</Badge>}
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: '180px 1fr', rowGap: 4, fontSize: 11 }}>
        <span style={{ color: tokens.colorNeutralForeground3 }}>Display name</span>
        <span>{col.displayName}</span>
        <span style={{ color: tokens.colorNeutralForeground3 }}>Thumbnail size</span>
        <span>144 × 144 (auto-cropped center)</span>
        <span style={{ color: tokens.colorNeutralForeground3 }}>Full-size</span>
        <span>{col.canStoreFullImage ? `Up to ${formatSize(maxBytes)}` : <em style={{ color: tokens.colorNeutralForeground3 }}>not stored — thumbnail only</em>}</span>
        <span style={{ color: tokens.colorNeutralForeground3 }}>Companion columns</span>
        <span style={{ fontFamily: tokens.fontFamilyMonospace, fontSize: 10 }}>
          {(() => { const c = imageCompanionsFor(col.logicalName); return `${c.id} · ${c.timestamp} · ${c.url}`; })()}
        </span>
        <span style={{ color: tokens.colorNeutralForeground3 }}>Set on Create?</span>
        <span>{col.isPrimaryImage ? 'Yes — primary image only' : 'No — Update or PATCH only'}</span>
        <span style={{ color: tokens.colorNeutralForeground3 }}>File types</span>
        <span>image/gif · image/jpeg · image/png · image/bmp · image/tiff</span>
      </div>
    </div>
  );
}
