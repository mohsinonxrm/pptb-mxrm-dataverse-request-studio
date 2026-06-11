// State shapes for the Binary group (File / Image / Attachment / Annotation
// column + record operations).
//
// Manage File Data:
//   • Upload — three supported approaches:
//       (a) single-request PATCH /<set>(<id>)/<col>     · x-ms-file-name header · body=base64
//       (b) chunked PATCH                               · init returns Location + x-ms-chunk-size; each chunk PATCH with Content-Range
//       (c) Dataverse messages                          · InitializeFileBlocksUpload → UploadBlock × N → CommitFileBlocksUpload
//   • Download — three:
//       (a) GET /<set>(<id>)/<col>/$value
//       (b) GET with Range: bytes=A-B (chunked)
//       (c) Dataverse messages: InitializeFileBlocksDownload → DownloadBlock × N
//       (d) GetFileSasUrl function (alternative — returns SAS link valid for 1 hour)
//   • Delete — DELETE /<set>(<id>)/<col>
//
// Manage Image Data:
//   • Primary image (entityimage)   — set on Create, always thumbnail 144×144.
//   • Custom image columns          — full-size optional via CanStoreFullImage.
//   • Upload / Download / Delete    — same patterns as File but with image-specific size flag.
//
// Manage Attachment / Annotation Data (attachment-annotation-files doc):
//   • Attachment (activitymimeattachment.body)        — must associate with email/template
//   • Annotation (annotation.documentbody)            — caller-generated annotationid required
//   • Upload — two supported approaches:
//       (a) inline base64 in create/update — for files ≲ 4 MB
//       (b) Dataverse messages:
//             Initialize{Attachment|Annotation}BlocksUpload  → UploadBlock × N
//             Commit{Attachment|Annotation}BlocksUpload      → returns FileSizeInBytes + record id
//   • Download — two:
//       (a) GET /<set>(<id>)/{body|documentbody}/$value      (returns base64 text/plain)
//       (b) Dataverse messages:
//             Initialize{Attachment|Annotation}BlocksDownload → DownloadBlock × N
//   • File size limit is org-wide via Organization.MaxUploadFileSize (default 5 MB, max 128 MB).
//
// Reference docs (all webapi tab):
//   files-images-overview
//   file-column-data
//   image-column-data
//   attachment-annotation-files
//   getfilesasurl

import type { HeaderItem } from '../editors/HeadersEditor';

// ──────────────────────────────────────────────────────────────
// Manage File Data
// ──────────────────────────────────────────────────────────────
export type FileOperation = 'upload' | 'download' | 'delete';

export type FileUploadMethod =
  | { kind: 'single-request' }
  | { kind: 'chunked-patch'; chunkSize: number } // chunkSize default = 2 MB (server may set via x-ms-chunk-size)
  | { kind: 'dataverse-messages'; blockSize: number }; // blockSize default = 4 MB

export type FileDownloadMethod =
  | { kind: 'single-request' }
  | { kind: 'ranged'; rangeStart: number; rangeEnd: number }
  | { kind: 'dataverse-messages'; blockSize: number }
  | { kind: 'sas-url' };

/** Source of the binary payload during upload. */
export type BinarySource =
  | { kind: 'file' } // user picks a file via <input type="file"> (we read into memory)
  | { kind: 'base64' } // user pastes a base64 string
  | { kind: 'url' }; // informational — references an external URL

export interface ManageFileState {
  /** Source table logical name. */
  table: string;
  /** Source row GUID. */
  recordId: string | null;
  /** Logical name of the File-typed column. */
  fileColumn: string | null;
  operation: FileOperation;
  uploadMethod: FileUploadMethod;
  downloadMethod: FileDownloadMethod;

  /** Upload payload metadata. */
  source: BinarySource;
  fileName: string;
  fileSize: number;
  mimeType: string;
  bodyBase64: string;
  /** When source.kind === 'url' — informational only; not transmitted. */
  bodyUrl: string;

  headers: HeaderItem[];
  dirty: Set<string>;
}

// ──────────────────────────────────────────────────────────────
// Manage Image Data
// ──────────────────────────────────────────────────────────────
export interface ManageImageState {
  table: string;
  recordId: string | null;
  imageColumn: string | null;
  operation: FileOperation;

  /**
   * Download size — thumbnail (144×144, default) vs full (?size=full,
   * only if column's canStoreFullImage is true).
   */
  downloadSize: 'thumbnail' | 'full';

  source: BinarySource;
  fileName: string;
  bodyBase64: string;
  bodyUrl: string;
  mimeType: 'image/png' | 'image/jpeg' | 'image/gif' | 'image/bmp' | 'image/tiff';

  headers: HeaderItem[];
  dirty: Set<string>;
}

// ──────────────────────────────────────────────────────────────
// Manage Attachment / Annotation Data
// ──────────────────────────────────────────────────────────────

/**
 * Discriminator between the two attachment-shaped tables. Both follow the same
 * message pattern; only the table name, body-column name, and message names
 * differ. We keep them in a single state so the mode UI can flip between them
 * without losing payload.
 */
export type AttachmentTarget = 'attachment' | 'annotation';

/** Per-target metadata derived from the discriminator — kept in code, not state. */
export const ATTACHMENT_TARGET_INFO: Record<
  AttachmentTarget,
  {
    /** Table logical name. */
    entityLogical: string;
    /** Entity set name (plural). */
    entitySet: string;
    /** Primary-key attribute (also the "target id" prop in Initialize* messages). */
    primaryKey: string;
    /** Body column logical name — base64-encoded string payload. */
    bodyColumn: string;
    /** File name column logical name (always 'filename' for both). */
    fileNameColumn: string;
    /** MIME type column logical name (always 'mimetype' for both). */
    mimeTypeColumn: string;
    /** Init upload message name. */
    initUploadMessage: string;
    /** Commit upload message name. */
    commitUploadMessage: string;
    /** Init download message name. */
    initDownloadMessage: string;
    /** Caller-supplied id required (annotation) vs server-generated (attachment + email parent). */
    callerSuppliesId: boolean;
  }
> = {
  attachment: {
    entityLogical: 'activitymimeattachment',
    entitySet: 'activitymimeattachments',
    primaryKey: 'activitymimeattachmentid',
    bodyColumn: 'body',
    fileNameColumn: 'filename',
    mimeTypeColumn: 'mimetype',
    initUploadMessage: 'InitializeAttachmentBlocksUpload',
    commitUploadMessage: 'CommitAttachmentBlocksUpload',
    initDownloadMessage: 'InitializeAttachmentBlocksDownload',
    callerSuppliesId: false,
  },
  annotation: {
    entityLogical: 'annotation',
    entitySet: 'annotations',
    primaryKey: 'annotationid',
    bodyColumn: 'documentbody',
    fileNameColumn: 'filename',
    mimeTypeColumn: 'mimetype',
    initUploadMessage: 'InitializeAnnotationBlocksUpload',
    commitUploadMessage: 'CommitAnnotationBlocksUpload',
    initDownloadMessage: 'InitializeAnnotationBlocksDownload',
    callerSuppliesId: true,
  },
};

export type AttachmentUploadMethod =
  | { kind: 'inline-base64' } // PATCH/POST with body column = base64 (≲4 MB)
  | { kind: 'dataverse-messages'; blockSize: number }; // blockSize default = 4 MB

export type AttachmentDownloadMethod =
  | { kind: 'single-request' } // GET /<set>(<id>)/{body|documentbody}/$value
  | { kind: 'dataverse-messages'; blockSize: number };

export interface ManageAttachmentState {
  /** Which record type — drives table, body column, and message names. */
  target: AttachmentTarget;

  /**
   * For 'attachment' on upload: the record will be created; activityId comes
   * from `parentActivityId` and is bound via `objectid_email@odata.bind`.
   * For 'annotation' on upload: caller-generated annotation GUID set via `recordId`.
   * For download/delete: the existing attachment/annotation GUID set via `recordId`.
   */
  recordId: string | null;

  /**
   * Attachment-only: the email activity GUID this attachment associates with.
   * Bound via `objectid_email@odata.bind: emails(<activityId>)` on Initialize*Upload.
   */
  parentActivityId: string | null;

  /**
   * Annotation-only: optional parent record reference (objectid). Annotations
   * can hang off any entity with HasNotes=true. Bound via
   * `objectid_<entitylogical>@odata.bind: <set>(<id>)` on the annotation Target.
   */
  parentEntityLogical: string | null;
  parentEntityId: string | null;

  /** Operation type. */
  operation: FileOperation;

  /** Upload / download method selectors. */
  uploadMethod: AttachmentUploadMethod;
  downloadMethod: AttachmentDownloadMethod;

  /** Upload payload metadata. */
  source: BinarySource;
  fileName: string;
  fileSize: number;
  mimeType: string;
  bodyBase64: string;
  /** When source.kind === 'url' — informational only; not transmitted. */
  bodyUrl: string;

  /** Attachment-only: free-text subject the email attachment carries. */
  subject: string;
  /** Annotation-only: notetext column (the note's body text — separate from documentbody). */
  noteText: string;

  headers: HeaderItem[];
  dirty: Set<string>;
}

// ──────────────────────────────────────────────────────────────
// Pre-flight size advisories (shared helpers)
// ──────────────────────────────────────────────────────────────

/**
 * Output of a size pre-flight check — surfaced in the pipeline card so the
 * user sees the verdict BEFORE attempting an upload that the server will
 * reject. Sourced from FileAttributeMetadata.MaxSizeInKB,
 * ImageAttributeMetadata.MaxSizeInKB (cap 30 MB), or
 * Organization.MaxUploadFileSize (attachment / annotation; default 5 MB,
 * max 128 MB).
 */
export interface SizeAdvisory {
  level: 'ok' | 'warn' | 'error';
  /** Short headline (e.g. "File exceeds column limit"). */
  headline: string;
  /** Detail copy describing the threshold + observed size. */
  detail: string;
}

/**
 * Compute a SizeAdvisory for a given upload size vs a column / org cap.
 *
 *   `maxBytes` — the cap, in bytes, derived from MaxSizeInKB * 1024 (for File /
 *                Image columns) or Organization.MaxUploadFileSize (for
 *                attachment / annotation). When undefined we surface a softer
 *                "size unknown" warning so the user knows we couldn't validate.
 *   `fileSize` — observed payload size (0 when no payload yet).
 *   `kind`     — what's being uploaded — drives the headline copy.
 */
export function checkUploadSize(
  fileSize: number,
  maxBytes: number | undefined,
  kind: 'file' | 'image' | 'attachment' | 'annotation',
): SizeAdvisory | null {
  if (fileSize <= 0) return null; // No payload yet — no advisory.
  const kindLabel =
    kind === 'file'
      ? 'file column'
      : kind === 'image'
        ? 'image column'
        : kind === 'attachment'
          ? 'attachment'
          : 'annotation';
  const fmt = (n: number) => formatBytes(n);
  if (maxBytes === undefined) {
    return {
      level: 'warn',
      headline: 'Size limit unknown',
      detail: `Could not resolve MaxSizeInKB for this ${kindLabel}. Pre-flight check skipped — ${fmt(fileSize)} payload will be sent; the server enforces its own cap.`,
    };
  }
  if (fileSize > maxBytes) {
    const overBy = fileSize - maxBytes;
    return {
      level: 'error',
      headline: `Exceeds ${kindLabel} limit by ${fmt(overBy)}`,
      detail: `Payload is ${fmt(fileSize)} but the configured cap is ${fmt(maxBytes)}. The server will return ProcessImageFailure / unManagedidsattachmentinvalidfilesize / similar. Reduce the payload or raise the limit (column MaxSizeInKB for file/image; Organization.MaxUploadFileSize for attachment/annotation, max 128 MB).`,
    };
  }
  if (fileSize > maxBytes * 0.9) {
    return {
      level: 'warn',
      headline: 'Approaching size limit',
      detail: `Payload is ${fmt(fileSize)} — ${Math.round((fileSize / maxBytes) * 100)}% of the ${fmt(maxBytes)} ${kindLabel} cap. Will succeed, but is close enough to fail if the source file grows.`,
    };
  }
  return {
    level: 'ok',
    headline: 'Size OK',
    detail: `${fmt(fileSize)} / ${fmt(maxBytes)} ${kindLabel} cap (${Math.round((fileSize / maxBytes) * 100)}% utilised).`,
  };
}

/** Internal — small KB/MB formatter (no decimals when under 10 of a unit). */
function formatBytes(bytes: number): string {
  if (!Number.isFinite(bytes) || bytes <= 0) return '0 B';
  const units = ['B', 'KB', 'MB', 'GB'];
  let v = bytes;
  let u = 0;
  while (v >= 1024 && u < units.length - 1) {
    v /= 1024;
    u++;
  }
  return `${v.toFixed(v < 10 && u > 0 ? 1 : 0)} ${units[u]}`;
}

// ──────────────────────────────────────────────────────────────
// Companion-column helpers
// ──────────────────────────────────────────────────────────────

/**
 * For a File-typed column, returns the auto-managed `<col>_Name` companion
 * column that stores the original filename. Useful in $select for retrieve
 * patterns so the caller knows what to name the downloaded blob without
 * actually transferring its bytes.
 */
export function fileNameColumnFor(fileColumnLogicalName: string): string {
  return `${fileColumnLogicalName}_Name`;
}

/**
 * For an Image-typed column, returns the auto-managed companion-column tuple:
 *   `<col>Id`         — image id (Guid)
 *   `<col>_Timestamp` — last-updated epoch (BigInt)
 *   `<col>_URL`       — relative download URL (string)
 *
 * Per image-column-data doc, none of these appear in the Power Apps designer
 * but all are queryable via the Web API.
 */
export function imageCompanionsFor(imageColumnLogicalName: string): {
  id: string;
  timestamp: string;
  url: string;
} {
  return {
    id: `${imageColumnLogicalName}Id`,
    timestamp: `${imageColumnLogicalName}_Timestamp`,
    url: `${imageColumnLogicalName}_URL`,
  };
}
