// PPTB forwards Dataverse calls over Electron IPC. When the main process
// rejects, Electron rewraps the error as
//   "Error invoking remote method 'dataverse.update': Error: <host message>"
// so the user sees plumbing before the Dataverse code and message (#188).
// Strip that wrapper (and the stringified "Error: " after it) for display;
// the executor keeps the raw text in the result body for Copy / JSON view.

const IPC_WRAPPER_RE = /Error invoking remote method '[^']*':\s*(?:Error:\s*)?/g;

/** The host message without Electron's IPC wrapper. */
export function cleanHostMessage(message: string): string {
  const cleaned = message.replace(IPC_WRAPPER_RE, '').trim();
  return cleaned || message;
}
