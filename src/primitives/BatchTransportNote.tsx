// The note shown where a request shape that Execute sends as a one-part
// `$batch` is chosen (#225): PUT of one column, clearing one column, and
// Upsert by alternate key. ToolBox's own impersonation doesn't reach a
// `$batch` part, so the note points at the Impersonation header instead.

import type { CSSProperties } from 'react';
import { MessageBar, MessageBarBody } from '@fluentui/react-components';

/** The note's text, for the modes' tests. */
export const BATCH_TRANSPORT_NOTE =
  "Sent through Power Platform ToolBox's $batch, so ToolBox's own impersonation setting " +
  "doesn't apply. To run it as another user, turn on MSCRMCallerID or CallerObjectId in HTTP headers.";

export function BatchTransportNote({ style }: { style?: CSSProperties }) {
  return (
    <MessageBar layout="multiline" intent="info" style={style}>
      <MessageBarBody>{BATCH_TRANSPORT_NOTE}</MessageBarBody>
    </MessageBar>
  );
}
