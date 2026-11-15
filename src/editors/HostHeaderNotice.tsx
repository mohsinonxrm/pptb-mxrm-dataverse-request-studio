// HostHeaderNotice — inline note for panes whose output is a request header
// (Precondition, Bypass). Power Platform ToolBox sends the header once the
// user approves it in the host, so this is an info note about that prompt.
// Standalone, nothing is sent, so there's nothing to note.

import { MessageBar, MessageBarBody, MessageBarTitle } from '@fluentui/react-components';
import type { ReactNode } from 'react';
import { useHostSession } from '../host/HostContext';

export function HostHeaderNotice({
  header,
}: {
  /** The header(s) this pane emits, for the pass-through note. */
  header: ReactNode;
}) {
  const host = useHostSession();
  if (!host.embedded) return null;
  return (
    <MessageBar layout="multiline" intent="info" style={{ marginBottom: 14 }}>
      <MessageBarBody>
        <MessageBarTitle>Sent by Power Platform ToolBox after you approve it</MessageBarTitle>
        Power Platform ToolBox sends {header} with the request. If you haven&apos;t approved
        additional request headers for this tool yet, it asks you first. Choose{' '}
        <strong>Approve for this tool</strong> so it doesn&apos;t ask again. If you reject, the
        request isn&apos;t sent. You can revoke an approval in ToolBox&apos;s Consent Review.
      </MessageBarBody>
    </MessageBar>
  );
}
