// Advice for an error the PPTB host reported without an HTTP status (#209).

/** A Dataverse error code, e.g. `0x80060882`. */
export const DATAVERSE_CODE_RE = /\b0x8[0-9a-f]{7}\b/i;

// Messages the host raises itself (desktop-app dataverseManager): its
// response-size limit, network failures, connection / token problems and
// argument checks. They never reached Dataverse, or never came back from it.
const HOST_MESSAGE_RE =
  /exceeds the maximum size|Request failed:|Connection .*not found|No access token|cannot be empty|Failed to (get|refresh|acquire)/i;

/**
 * Where a no-status error came from: Dataverse when it carries a Dataverse
 * code, the host for its own known messages, and neither claimed otherwise
 * (Dataverse sometimes sends an error with an empty code).
 */
export function noStatusAdvice(err: { code?: string; message?: string } | undefined): string {
  const text = `${err?.code ?? ''} ${err?.message ?? ''}`;
  if (DATAVERSE_CODE_RE.test(text)) {
    return "The PPTB host didn't report an HTTP status. The error message below comes from Dataverse.";
  }
  if (HOST_MESSAGE_RE.test(text)) {
    return "The PPTB host didn't report an HTTP status. The message below comes from the host itself, not from Dataverse.";
  }
  return "The PPTB host didn't report an HTTP status. Inspect the message below for details.";
}
