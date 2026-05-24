// URL + body construction for the Binary group (File / Image columns).
//
// A single mode call may produce 1+ HTTP requests (init → upload chunks →
// commit). Each builder returns either a single BuiltRequest (for the URL
// bar) or a full pipeline of requests for the Pipeline pane + multi-request
// code generation.
//
// Grounded in:
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/files-images-overview?tabs=webapi
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/file-column-data?tabs=webapi
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/image-column-data?tabs=webapi
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/attachment-annotation-files?tabs=webapi
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/getfilesasurl?tabs=webapi

import { ENV } from '../mock/environment';
import { findTable } from '../mock/metadata';
import type { BuiltRequest } from './urlBuilder';
import {
  ATTACHMENT_TARGET_INFO,
  type ManageFileState,
  type ManageImageState,
  type ManageAttachmentState,
  type AttachmentTarget,
} from '../state/binaryState';

// ──────────────────────────────────────────────────────────────
// One step in a binary pipeline (init → upload-N → commit).
// Multi-step ops emit one BinaryPipelineStep per HTTP request so the
// Pipeline pane can render the full sequence with status / progress.
// ──────────────────────────────────────────────────────────────
export interface BinaryPipelineStep {
  /** Display index — 1, 2, 3 … */
  n: number;
  /** Human-friendly title — e.g. "Initialize upload". */
  title: string;
  /** HTTP method. */
  method: 'GET' | 'POST' | 'PATCH' | 'PUT' | 'DELETE';
  /** Path + query (no host). */
  relativeUrl: string;
  /** Per-request headers to call out in the UI (Content-Range etc.). */
  headers: Record<string, string>;
  /** Body shape — JSON object for action calls, opaque marker for binary chunks. */
  body?: Record<string, unknown> | '<base64 file bytes>' | '<chunk bytes>';
  /** Inline detail / description. */
  detail?: string;
}

// ──────────────────────────────────────────────────────────────
// File — single-request upload
//
//   PATCH /<set>(<id>)/<col>
//   x-ms-file-name: <name>                 (HEADER only — no query param)
//   Content-Type: application/octet-stream
//   <body: raw bytes (we surface as <base64 file bytes>)>
//   → 204 NoContent
// ──────────────────────────────────────────────────────────────
export function buildFileSingleUpload(s: ManageFileState): BinaryPipelineStep[] {
  const segment = fileSegment(s);
  if (!segment) return [];
  return [{
    n: 1,
    title: 'Upload file (single request)',
    method: 'PATCH',
    relativeUrl: segment,
    headers: {
      'x-ms-file-name': s.fileName || 'file.bin',
      'Content-Type': 'application/octet-stream',
    },
    body: '<base64 file bytes>',
    detail: `Single request — server writes ${s.fileName || 'file.bin'} (${formatSize(s.fileSize)}) directly and responds 204 NoContent. Doc-recommended for files up to 128 MB; larger files require chunked upload or the messages path.`,
  }];
}

// ──────────────────────────────────────────────────────────────
// File — chunked PATCH upload (per file-column-data docs)
//
//   1. PATCH /<set>(<id>)/<col>?x-ms-file-name=<name>   (filename in QUERY param)
//      x-ms-transfer-mode: chunked
//      → 200 OK · Location: <upload-url with sessiontoken=...>
//                · x-ms-chunk-size: <bytes>
//   2..N. PATCH <upload-url>                            (Location URL re-used per chunk)
//      x-ms-file-name: <name>                            (header on EVERY chunk)
//      Content-Type: application/octet-stream
//      Content-Range: bytes A-B/Total
//      Content-Length: <chunk-bytes>
//      < chunk bytes >
//      → 206 PartialContent for intermediate chunks · 204 NoContent on final
// ──────────────────────────────────────────────────────────────
export function buildFileChunkedUpload(s: ManageFileState): BinaryPipelineStep[] {
  const segment = fileSegment(s);
  if (!segment || s.uploadMethod.kind !== 'chunked-patch') return [];
  const chunkSize = s.uploadMethod.chunkSize;
  const total = Math.max(s.fileSize, 1);
  const chunkCount = Math.max(1, Math.ceil(total / chunkSize));
  const fileName = s.fileName || 'file.bin';

  const steps: BinaryPipelineStep[] = [{
    n: 1,
    title: 'Initialize chunked upload',
    method: 'PATCH',
    relativeUrl: `${segment}?x-ms-file-name=${encodeURIComponent(fileName)}`,
    headers: {
      'x-ms-transfer-mode': 'chunked',
    },
    detail: `Filename travels as a query param on init; the only request header is x-ms-transfer-mode. Response 200 → Location: <upload-url with sessiontoken=…> · x-ms-chunk-size: ${chunkSize.toLocaleString()} bytes (≈ ${formatSize(chunkSize)}). Server may override the chunk size — honor x-ms-chunk-size in the response.`,
  }];

  // Surface the FIRST chunk PATCH inline so the user sees the shape; the rest
  // are summarized as "N more chunks". We don't emit thousands of steps.
  const firstByteStart = 0;
  const firstByteEnd = Math.min(chunkSize, total) - 1;
  const firstChunkBytes = firstByteEnd - firstByteStart + 1;
  steps.push({
    n: 2,
    title: `Upload chunk 1 of ${chunkCount}`,
    method: 'PATCH',
    relativeUrl: '<upload-url-from-Location-header>',
    headers: {
      'x-ms-file-name': fileName,
      'Content-Type': 'application/octet-stream',
      'Content-Range': `bytes ${firstByteStart}-${firstByteEnd}/${total}`,
      'Content-Length': String(firstChunkBytes),
    },
    body: '<chunk bytes>',
    detail: chunkCount === 1
      ? 'Single chunk — server commits on receipt and returns 204 NoContent.'
      : `Intermediate chunks return 206 PartialContent. Repeat with Content-Range stepping by ${chunkSize.toLocaleString()} bytes per chunk; the final chunk's Content-Length may be smaller.`,
  });

  if (chunkCount > 1) {
    const lastChunkBytes = total - (chunkCount - 1) * chunkSize;
    steps.push({
      n: 3,
      title: `${chunkCount - 1} more chunk${chunkCount - 1 === 1 ? '' : 's'}`,
      method: 'PATCH',
      relativeUrl: '<upload-url-from-Location-header>',
      headers: {
        'x-ms-file-name': fileName,
        'Content-Type': 'application/octet-stream',
        'Content-Range': `bytes <start>-<end>/${total}`,
        'Content-Length': '<chunk-size>',
      },
      body: '<chunk bytes>',
      detail: `Final chunk Content-Range ends at ${(total - 1).toLocaleString()} with Content-Length ${lastChunkBytes.toLocaleString()}. Server returns 204 NoContent on the final chunk (commits automatically).`,
    });
  }
  return steps;
}

// ──────────────────────────────────────────────────────────────
// File — Dataverse messages (Initialize/Upload/Commit)
//
//   POST /InitializeFileBlocksUpload  →  FileContinuationToken
//   POST /UploadBlock                 ×  N  (BlockId + BlockData + Token)
//   POST /CommitFileBlocksUpload      →  FileId
// ──────────────────────────────────────────────────────────────
export function buildFileMessagesUpload(s: ManageFileState): BinaryPipelineStep[] {
  if (s.uploadMethod.kind !== 'dataverse-messages') return [];
  const tbl = findTable(s.table);
  if (!tbl) return [];
  const blockSize = s.uploadMethod.blockSize;
  const total = Math.max(s.fileSize, 1);
  const blockCount = Math.max(1, Math.ceil(total / blockSize));
  const targetRef = {
    '@odata.type': `Microsoft.Dynamics.CRM.${tbl.logicalName}`,
    [tbl.primaryKey]: s.recordId ?? '<record-id>',
  };
  return [
    {
      n: 1,
      title: 'InitializeFileBlocksUpload',
      method: 'POST',
      relativeUrl: `${ENV.apiBase}/InitializeFileBlocksUpload`,
      headers: { 'Content-Type': 'application/json; charset=utf-8' },
      body: {
        Target: targetRef,
        FileAttributeName: s.fileColumn ?? '<file-column>',
        FileName: s.fileName || 'file.bin',
      },
      detail: 'Returns FileContinuationToken used by subsequent UploadBlock / CommitFileBlocksUpload calls.',
    },
    {
      n: 2,
      title: `UploadBlock × ${blockCount}`,
      method: 'POST',
      relativeUrl: `${ENV.apiBase}/UploadBlock`,
      headers: { 'Content-Type': 'application/json; charset=utf-8' },
      body: {
        BlockId: '<base64 blockid — same length for every block>',
        BlockData: '<base64 chunk bytes (≤ 4 MB)>',
        FileContinuationToken: '<continuation-token-from-step-1>',
      },
      detail: `${blockCount} request${blockCount === 1 ? '' : 's'} of ${formatSize(blockSize)} each. Capture BlockId values for CommitFileBlocksUpload.`,
    },
    {
      n: 3,
      title: 'CommitFileBlocksUpload',
      method: 'POST',
      relativeUrl: `${ENV.apiBase}/CommitFileBlocksUpload`,
      headers: { 'Content-Type': 'application/json; charset=utf-8' },
      body: {
        BlockList: ['<blockid-1>', '<blockid-2>', '…'],
        FileContinuationToken: '<continuation-token-from-step-1>',
        FileName: s.fileName || 'file.bin',
        MimeType: s.mimeType || 'application/octet-stream',
      },
      detail: 'Returns FileId (Edm.Guid) — store this for later DeleteFile or download by id.',
    },
  ];
}

// ──────────────────────────────────────────────────────────────
// File — single-request download
//
//   GET /<set>(<id>)/<col>/$value
//   (optionally with Range: bytes=A-B for partial)
// ──────────────────────────────────────────────────────────────
export function buildFileDownload(s: ManageFileState): BinaryPipelineStep[] {
  const segment = fileSegment(s);
  if (!segment) return [];
  switch (s.downloadMethod.kind) {
    case 'single-request':
      return [{
        n: 1, title: 'Download file', method: 'GET',
        relativeUrl: `${segment}/$value`,
        headers: { Accept: 'application/octet-stream' },
        detail: `Response 200 OK · binary body · response headers carry x-ms-file-size, x-ms-file-name, mimetype. To know the filename without transferring bytes first, $select the auto-managed companion column: GET /<set>(<id>)?$select=${s.fileColumn ?? '<col>'}_Name. Returns the file in a single response — for >128 MB prefer ranged download or the messages path.`,
      }];
    case 'ranged': {
      const r = s.downloadMethod;
      return [{
        n: 1, title: 'Download (Range)', method: 'GET',
        relativeUrl: `${segment}/$value`,
        headers: {
          Accept: 'application/octet-stream',
          Range: `bytes=${r.rangeStart}-${r.rangeEnd}`,
        },
        detail: 'Response 206 PartialContent · response headers carry x-ms-file-size (total), x-ms-file-name, x-ms-chunk-size, mimetype. Use x-ms-file-size to drive subsequent Range requests until the whole file is downloaded.',
      }];
    }
    case 'dataverse-messages': {
      const tbl = findTable(s.table);
      const targetRef = tbl ? {
        '@odata.type': `Microsoft.Dynamics.CRM.${tbl.logicalName}`,
        [tbl.primaryKey]: s.recordId ?? '<record-id>',
      } : {};
      return [
        {
          n: 1, title: 'InitializeFileBlocksDownload', method: 'POST',
          relativeUrl: `${ENV.apiBase}/InitializeFileBlocksDownload`,
          headers: { 'Content-Type': 'application/json; charset=utf-8' },
          body: { Target: targetRef, FileAttributeName: s.fileColumn ?? '<file-column>' },
          detail: 'Returns { FileContinuationToken, FileSizeInBytes, FileName, IsChunkingSupported }. If IsChunkingSupported is false (database-stored files), BlockLength on DownloadBlock must equal FileSizeInBytes — only one call.',
        },
        {
          n: 2, title: 'DownloadBlock × N', method: 'POST',
          relativeUrl: `${ENV.apiBase}/DownloadBlock`,
          headers: { 'Content-Type': 'application/json; charset=utf-8' },
          body: {
            Offset: 0,
            BlockLength: s.downloadMethod.blockSize,
            FileContinuationToken: '<token-from-step-1>',
          },
          detail: 'Returns { Data: <base64 chunk bytes> }. Increment Offset by BlockLength on each call. Final block may be smaller; BlockLength can stay constant. For image columns this path always returns the FULL-size image — thumbnail-only retrieval requires the $value endpoint instead.',
        },
      ];
    }
    case 'sas-url': {
      const tbl = findTable(s.table);
      const setName = tbl?.entitySetName ?? '<set>';
      const id = s.recordId ?? '<record-id>';
      return [{
        n: 1, title: 'GetFileSasUrl', method: 'GET',
        relativeUrl: `${ENV.apiBase}/GetFileSasUrl(Target=@t,FileAttributeName=@f)?@t={'@odata.id':'${setName}(${id})'}&@f='${s.fileColumn ?? '<col>'}'`,
        headers: { Accept: 'application/json' },
        detail: 'Returns FileName + FileSizeInBytes + MimeType + SasUrl (valid 1 hour). Anyone with the URL can download. Requires full-size image support for image columns.',
      }];
    }
  }
}

// ──────────────────────────────────────────────────────────────
// File — delete
//
//   DELETE /<set>(<id>)/<col>                    → 204 NoContent
//   alt:   POST /DeleteFile { FileId: "guid" }   → 204 NoContent
// ──────────────────────────────────────────────────────────────
export function buildFileDelete(s: ManageFileState): BinaryPipelineStep[] {
  const segment = fileSegment(s);
  if (!segment) return [];
  return [{
    n: 1, title: 'Delete file column value', method: 'DELETE',
    relativeUrl: segment,
    headers: { Accept: 'application/json' },
    detail: 'Response 204 NoContent — clears the column. Alternative form (when you only have the FileId): POST /DeleteFile { "FileId": "<guid>" }. File columns can NOT be cleared via PATCH-to-null — that path is image-only.',
  }];
}

// ──────────────────────────────────────────────────────────────
// Image — upload (PATCH set body = base64)
//
//   PATCH /<set>(<id>)
//   If-Match: *
//   { "<col>": "<base64 string>" }              → 204 NoContent
//
// Per image-column-data doc this is the canonical thumbnail-update path.
// For columns where canStoreFullImage=true and the user wants to upload the
// FULL-size original, the same chunked-PATCH / Dataverse-messages APIs used
// for file columns are required ("To upload images, use the same APIs you
// use to upload files" — image-column-data doc). When CanStoreFullImage is
// false, this PATCH is the only supported upload path.
// ──────────────────────────────────────────────────────────────
export function buildImageUpload(s: ManageImageState): BinaryPipelineStep[] {
  const tbl = findTable(s.table);
  if (!tbl) return [];
  const idSeg = s.recordId ? `(${s.recordId})` : '(<record-id>)';
  return [{
    n: 1, title: 'Upload image', method: 'PATCH',
    relativeUrl: `${ENV.apiBase}/${tbl.entitySetName}${idSeg}`,
    headers: {
      'If-Match': '*',
      'Content-Type': 'application/json; charset=utf-8',
    },
    body: {
      [s.imageColumn ?? '<image-column>']: '<base64 image bytes>',
    },
    detail: 'Response 204 NoContent. Doc-canonical thumbnail upload — image goes in the JSON body as a base 64 string. For full-size uploads on columns with CanStoreFullImage=true, use the file-column chunked-PATCH or InitializeFileBlocksUpload / UploadBlock / CommitFileBlocksUpload sequence (same APIs as file columns).',
  }];
}

// ──────────────────────────────────────────────────────────────
// Image — download
//
//   GET /<set>(<id>)/<col>/$value           (thumbnail, default)
//   GET /<set>(<id>)/<col>/$value?size=full (full, only if canStoreFullImage)
// ──────────────────────────────────────────────────────────────
export function buildImageDownload(s: ManageImageState): BinaryPipelineStep[] {
  const tbl = findTable(s.table);
  if (!tbl) return [];
  const idSeg = s.recordId ? `(${s.recordId})` : '(<record-id>)';
  const sizeParam = s.downloadSize === 'full' ? '?size=full' : '';
  return [{
    n: 1, title: s.downloadSize === 'full' ? 'Download full-size image' : 'Download thumbnail', method: 'GET',
    relativeUrl: `${ENV.apiBase}/${tbl.entitySetName}${idSeg}/${s.imageColumn ?? '<image-column>'}/$value${sizeParam}`,
    headers: { Accept: 'image/*' },
    detail: s.downloadSize === 'full'
      ? 'Response 200 OK + binary image body when CanStoreFullImage=true. When CanStoreFullImage=false (or the image was uploaded before the flag was enabled), the server returns 204 NoContent with an empty body — not an error.'
      : 'Response 200 OK + 144×144 thumbnail. Thumbnails are cropped/resized server-side: images ≥144 px on either side are cropped on-center to 144×144; images smaller than 144 px on both sides are cropped square to the smallest side. Append ?size=full for the original (requires CanStoreFullImage=true).',
  }];
}

// ──────────────────────────────────────────────────────────────
// Image — delete (3 doc-equivalent forms — all return 204 NoContent)
//
//   (a) DELETE /<set>(<id>)/<col>                          ← emitted as the canonical form
//   (b) PATCH  /<set>(<id>)            { <col>: null }
//   (c) PUT    /<set>(<id>)/<col>      { "value": null }
//
// Per image-column-data doc, images (unlike file columns) CAN be cleared via
// PATCH-to-null because they're carried inline on the record.
// ──────────────────────────────────────────────────────────────
export function buildImageDelete(s: ManageImageState): BinaryPipelineStep[] {
  const tbl = findTable(s.table);
  if (!tbl) return [];
  const idSeg = s.recordId ? `(${s.recordId})` : '(<record-id>)';
  return [{
    n: 1, title: 'Delete image (DELETE)', method: 'DELETE',
    relativeUrl: `${ENV.apiBase}/${tbl.entitySetName}${idSeg}/${s.imageColumn ?? '<image-column>'}`,
    headers: { Accept: 'application/json' },
    detail: 'Response 204 NoContent. Three doc-equivalent forms: (a) DELETE on the column URL (shown), (b) PATCH /set(id) with { "<col>": null }, (c) PUT /set(id)/<col> with { "value": null }. Unlike file columns, images can be cleared via PATCH-to-null since they ride inline on the record.',
  }];
}

// ──────────────────────────────────────────────────────────────
// Helpers
// ──────────────────────────────────────────────────────────────
function fileSegment(s: ManageFileState): string | null {
  const tbl = findTable(s.table);
  if (!tbl) return null;
  const idSeg = s.recordId ? `(${s.recordId})` : '(<record-id>)';
  const colSeg = s.fileColumn ?? '<file-column>';
  return `${ENV.apiBase}/${tbl.entitySetName}${idSeg}/${colSeg}`;
}

export function formatSize(bytes: number): string {
  if (!Number.isFinite(bytes) || bytes <= 0) return '0 B';
  const units = ['B', 'KB', 'MB', 'GB'];
  let v = bytes;
  let u = 0;
  while (v >= 1024 && u < units.length - 1) { v /= 1024; u++; }
  return `${v.toFixed(v < 10 && u > 0 ? 1 : 0)} ${units[u]}`;
}

// ──────────────────────────────────────────────────────────────
// Aggregate builders — produce BuiltRequest for URL bar + pipeline
// for code-generator multi-request paths.
// ──────────────────────────────────────────────────────────────

/** Returns the FIRST step's URL/method as a BuiltRequest for the URL bar. */
export function buildManageFile(s: ManageFileState): BuiltRequest {
  const tbl = findTable(s.table);
  if (!tbl) return { relativeUrl: '', relativeNoBase: '', bytes: 0, queryParts: [], entitySet: '', entityLogical: '' };
  const steps = manageFilePipeline(s);
  const first = steps[0];
  if (!first) {
    return {
      relativeUrl: '', relativeNoBase: '', bytes: 0, queryParts: [],
      entitySet: tbl.entitySetName, entityLogical: tbl.logicalName,
      recordId: s.recordId ?? undefined,
    };
  }
  const noBase = first.relativeUrl.replace(ENV.apiBase, '');
  return {
    relativeUrl: first.relativeUrl,
    relativeNoBase: noBase || first.relativeUrl,
    bytes: first.relativeUrl.length,
    queryParts: [],
    entitySet: tbl.entitySetName,
    entityLogical: tbl.logicalName,
    recordId: s.recordId ?? undefined,
  };
}

export function manageFilePipeline(s: ManageFileState): BinaryPipelineStep[] {
  switch (s.operation) {
    case 'upload':
      switch (s.uploadMethod.kind) {
        case 'single-request':       return buildFileSingleUpload(s);
        case 'chunked-patch':        return buildFileChunkedUpload(s);
        case 'dataverse-messages':   return buildFileMessagesUpload(s);
      }
      break;
    case 'download':                  return buildFileDownload(s);
    case 'delete':                    return buildFileDelete(s);
  }
}

export function buildManageImage(s: ManageImageState): BuiltRequest {
  const tbl = findTable(s.table);
  if (!tbl) return { relativeUrl: '', relativeNoBase: '', bytes: 0, queryParts: [], entitySet: '', entityLogical: '' };
  const steps = manageImagePipeline(s);
  const first = steps[0];
  if (!first) {
    return {
      relativeUrl: '', relativeNoBase: '', bytes: 0, queryParts: [],
      entitySet: tbl.entitySetName, entityLogical: tbl.logicalName,
      recordId: s.recordId ?? undefined,
    };
  }
  const noBase = first.relativeUrl.replace(ENV.apiBase, '');
  return {
    relativeUrl: first.relativeUrl,
    relativeNoBase: noBase || first.relativeUrl,
    bytes: first.relativeUrl.length,
    queryParts: [],
    entitySet: tbl.entitySetName,
    entityLogical: tbl.logicalName,
    recordId: s.recordId ?? undefined,
  };
}

export function manageImagePipeline(s: ManageImageState): BinaryPipelineStep[] {
  switch (s.operation) {
    case 'upload':   return buildImageUpload(s);
    case 'download': return buildImageDownload(s);
    case 'delete':   return buildImageDelete(s);
  }
}

// ──────────────────────────────────────────────────────────────────────────────
// Attachment / Annotation builders
// ──────────────────────────────────────────────────────────────────────────────
//
// Per attachment-annotation-files docs, attachments (activitymimeattachment)
// and annotations (annotation) share the same wire shape with three pivots:
//   • table name + entity set
//   • body column ('body' vs 'documentbody')
//   • message-name prefix (Attachment vs Annotation on Initialize* and Commit*)
//
// UploadBlock and DownloadBlock are SHARED — same messages as file columns.
//
// Two upload paths:
//   (a) inline base64 — create/update record with body column populated
//   (b) messages       — Initialize*Upload → UploadBlock × N → Commit*Upload
//
// Two download paths:
//   (a) single  — GET /<set>(<id>)/{body|documentbody}/$value (returns text/plain base64)
//   (b) messages — Initialize*Download → DownloadBlock × N

/**
 * Build the {@link AttachmentTarget}-shaped Target envelope used by every
 * Initialize* / Commit* message. The shape differs between create (no id)
 * and update (id required).
 */
function attachmentTargetRef(s: ManageAttachmentState, includeBodyMeta: boolean): Record<string, unknown> {
  const info = ATTACHMENT_TARGET_INFO[s.target];
  const ref: Record<string, unknown> = {
    '@odata.type': `Microsoft.Dynamics.CRM.${info.entityLogical}`,
  };
  // Annotation: caller-supplied id. Attachment: only present on update.
  if (s.recordId) {
    ref[info.primaryKey] = s.recordId;
  }
  // Attachment associates to an email activity via objectid_email@odata.bind.
  if (s.target === 'attachment' && s.parentActivityId) {
    ref['objectid_email@odata.bind'] = `emails(${s.parentActivityId})`;
    ref['objecttypecode'] = 'email';
  }
  // Annotation can optionally associate to any HasNotes=true parent.
  if (s.target === 'annotation' && s.parentEntityLogical && s.parentEntityId) {
    const navName = `objectid_${s.parentEntityLogical}`;
    const parentTbl = findTable(s.parentEntityLogical);
    if (parentTbl) {
      ref[`${navName}@odata.bind`] = `${parentTbl.entitySetName}(${s.parentEntityId})`;
    }
  }
  if (includeBodyMeta) {
    if (s.subject && s.target === 'attachment') ref['subject'] = s.subject;
    if (s.noteText && s.target === 'annotation') ref['notetext'] = s.noteText;
    if (s.fileName) ref[info.fileNameColumn] = s.fileName || 'file.bin';
    if (s.mimeType) ref[info.mimeTypeColumn] = s.mimeType || 'application/octet-stream';
  }
  return ref;
}

// ──────────────────────────────────────────────────────────────
// Attachment / Annotation — inline-base64 upload
// ──────────────────────────────────────────────────────────────
export function buildAttachmentInlineUpload(s: ManageAttachmentState): BinaryPipelineStep[] {
  const info = ATTACHMENT_TARGET_INFO[s.target];
  // Inline upload is a CREATE (annotation: PATCH with caller-supplied id; attachment: POST).
  // For annotation, the doc shows PATCH /<set>(<id>) with the caller-supplied id.
  // For attachment, the doc shows POST /<set>.
  const isCreate = s.target === 'attachment' || !s.recordId;
  const method = isCreate ? 'POST' : 'PATCH';
  const url = isCreate
    ? `${ENV.apiBase}/${info.entitySet}`
    : `${ENV.apiBase}/${info.entitySet}(${s.recordId})`;

  // Build inline body — everything that goes in the Target envelope PLUS the body column.
  const body: Record<string, unknown> = attachmentTargetRef(s, true);
  // Strip @odata.type from the body — that's only needed when nested inside a Target.
  delete body['@odata.type'];
  body[info.bodyColumn] = '<base64 file bytes>';

  const headers: Record<string, string> = {
    'Content-Type': 'application/json; charset=utf-8',
  };
  if (!isCreate) headers['If-Match'] = '*';

  return [{
    n: 1,
    title: isCreate ? `Create ${s.target} with inline body` : `Update ${s.target} body`,
    method,
    relativeUrl: url,
    headers,
    body,
    detail: isCreate
      ? `Response ${s.target === 'attachment' ? '204 NoContent + OData-EntityId header' : '204 NoContent'} on success. Base 64-encoded payload must fit under Organization.MaxUploadFileSize (default 5 MB, max 128 MB). For larger files use the messages path.`
      : 'Response 204 NoContent. Replaces the existing body. For files larger than ~4 MB use the messages path instead.',
  }];
}

// ──────────────────────────────────────────────────────────────
// Attachment / Annotation — Dataverse messages upload
// ──────────────────────────────────────────────────────────────
export function buildAttachmentMessagesUpload(s: ManageAttachmentState): BinaryPipelineStep[] {
  if (s.uploadMethod.kind !== 'dataverse-messages') return [];
  const info = ATTACHMENT_TARGET_INFO[s.target];
  const blockSize = s.uploadMethod.blockSize;
  const total = Math.max(s.fileSize, 1);
  const blockCount = Math.max(1, Math.ceil(total / blockSize));

  const targetEnvelope = attachmentTargetRef(s, true);
  // The doc explicitly REMOVES body/documentbody from the Target on Initialize.
  // The body is uploaded via UploadBlock and committed via Commit*. Keep filename/mimetype.

  const commitProp = s.target === 'attachment' ? 'ActivityMimeAttachmentId' : 'AnnotationId';

  return [
    {
      n: 1,
      title: info.initUploadMessage,
      method: 'POST',
      relativeUrl: `${ENV.apiBase}/${info.initUploadMessage}`,
      headers: { 'Content-Type': 'application/json; charset=utf-8' },
      body: { Target: targetEnvelope },
      detail: s.target === 'attachment'
        ? 'Returns { FileContinuationToken }. Must include objectid_email@odata.bind on Target — these messages only CREATE attachments; trying to update an existing one returns "record already exists".'
        : 'Returns { FileContinuationToken }. Annotation must carry a caller-generated annotationid GUID on Target — Dataverse will not auto-assign. The body/documentbody must NOT be set in Target (use UploadBlock instead).',
    },
    {
      n: 2,
      title: `UploadBlock × ${blockCount}`,
      method: 'POST',
      relativeUrl: `${ENV.apiBase}/UploadBlock`,
      headers: { 'Content-Type': 'application/json; charset=utf-8' },
      body: {
        BlockId: '<base64 blockid — same length for every block (≤ 64 bytes before encoding)>',
        BlockData: '<base64 chunk bytes (≤ 4 MB)>',
        FileContinuationToken: '<continuation-token-from-step-1>',
      },
      detail: `${blockCount} request${blockCount === 1 ? '' : 's'} of ${formatSize(blockSize)} each. Each call returns 204 NoContent. Capture every BlockId — they MUST be passed to Commit* in upload order. Standard idiom: BlockId = Base64(UTF8(Guid.NewGuid().ToString())).`,
    },
    {
      n: 3,
      title: info.commitUploadMessage,
      method: 'POST',
      relativeUrl: `${ENV.apiBase}/${info.commitUploadMessage}`,
      headers: { 'Content-Type': 'application/json; charset=utf-8' },
      body: {
        Target: targetEnvelope,
        BlockList: ['<blockid-1>', '<blockid-2>', '…'],
        FileContinuationToken: '<continuation-token-from-step-1>',
      },
      detail: `Returns { ${commitProp}, FileSizeInBytes }. The Target envelope here must repeat the same filename/mimetype values from step 1 (server uses these to finalize the record). For attachments, the bound parent activity and subject also repeat here.`,
    },
  ];
}

// ──────────────────────────────────────────────────────────────
// Attachment / Annotation — download
// ──────────────────────────────────────────────────────────────
export function buildAttachmentDownload(s: ManageAttachmentState): BinaryPipelineStep[] {
  const info = ATTACHMENT_TARGET_INFO[s.target];
  switch (s.downloadMethod.kind) {
    case 'single-request': {
      const id = s.recordId ?? '<record-id>';
      return [{
        n: 1,
        title: `Download ${s.target} body (single request)`,
        method: 'GET',
        relativeUrl: `${ENV.apiBase}/${info.entitySet}(${id})/${info.bodyColumn}/$value`,
        headers: { Accept: 'application/json' },
        detail: `Response 200 OK · Content-Type: text/plain · body is the BASE 64-encoded blob (not raw bytes). Unlike file columns, this path does NOT return x-ms-file-name / x-ms-file-size / mimetype headers — you must read those via a separate $select on filename/mimetype.`,
      }];
    }
    case 'dataverse-messages': {
      return [
        {
          n: 1,
          title: info.initDownloadMessage,
          method: 'POST',
          relativeUrl: `${ENV.apiBase}/${info.initDownloadMessage}`,
          headers: { 'Content-Type': 'application/json; charset=utf-8' },
          body: { Target: attachmentTargetRef(s, false) },
          detail: 'Returns { FileContinuationToken, FileSizeInBytes, FileName }. Same DownloadBlock action used by file columns is then called repeatedly.',
        },
        {
          n: 2,
          title: 'DownloadBlock × N',
          method: 'POST',
          relativeUrl: `${ENV.apiBase}/DownloadBlock`,
          headers: { 'Content-Type': 'application/json; charset=utf-8' },
          body: {
            Offset: 0,
            BlockLength: s.downloadMethod.blockSize,
            FileContinuationToken: '<token-from-step-1>',
          },
          detail: 'Returns { Data: <base64 chunk bytes> }. Increment Offset by BlockLength on each call. Per doc note: BlockLength can stay constant — the final block may be smaller and the server returns only the remaining bytes.',
        },
      ];
    }
  }
}

// ──────────────────────────────────────────────────────────────
// Attachment / Annotation — delete
// ──────────────────────────────────────────────────────────────
export function buildAttachmentDelete(s: ManageAttachmentState): BinaryPipelineStep[] {
  const info = ATTACHMENT_TARGET_INFO[s.target];
  const id = s.recordId ?? '<record-id>';
  return [{
    n: 1,
    title: `Delete ${s.target} record`,
    method: 'DELETE',
    relativeUrl: `${ENV.apiBase}/${info.entitySet}(${id})`,
    headers: { Accept: 'application/json' },
    detail: `Response 204 NoContent. Deletes the entire ${s.target} record (including the body). Per attachment-annotation-files doc there's no message form for delete — the record IS the file, so removing the row removes the data. To clear just the body without deleting the record, PATCH ${info.bodyColumn} = null.`,
  }];
}

// ──────────────────────────────────────────────────────────────
// Aggregate builders (BuiltRequest + pipeline)
// ──────────────────────────────────────────────────────────────
export function buildManageAttachment(s: ManageAttachmentState): BuiltRequest {
  const info = ATTACHMENT_TARGET_INFO[s.target];
  const steps = manageAttachmentPipeline(s);
  const first = steps[0];
  if (!first) {
    return {
      relativeUrl: '', relativeNoBase: '', bytes: 0, queryParts: [],
      entitySet: info.entitySet, entityLogical: info.entityLogical,
      recordId: s.recordId ?? undefined,
    };
  }
  const noBase = first.relativeUrl.replace(ENV.apiBase, '');
  return {
    relativeUrl: first.relativeUrl,
    relativeNoBase: noBase || first.relativeUrl,
    bytes: first.relativeUrl.length,
    queryParts: [],
    entitySet: info.entitySet,
    entityLogical: info.entityLogical,
    recordId: s.recordId ?? undefined,
  };
}

export function manageAttachmentPipeline(s: ManageAttachmentState): BinaryPipelineStep[] {
  switch (s.operation) {
    case 'upload':
      switch (s.uploadMethod.kind) {
        case 'inline-base64':      return buildAttachmentInlineUpload(s);
        case 'dataverse-messages': return buildAttachmentMessagesUpload(s);
      }
      break;
    case 'download':               return buildAttachmentDownload(s);
    case 'delete':                 return buildAttachmentDelete(s);
  }
}

// ──────────────────────────────────────────────────────────────
// AttachmentTarget label helper (consumed by SaveContext narratives + mode picker)
// ──────────────────────────────────────────────────────────────
export function attachmentTargetLabel(t: AttachmentTarget): string {
  return t === 'attachment' ? 'Attachment (ActivityMimeAttachment)' : 'Note (Annotation)';
}
