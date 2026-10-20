// HostHeaderNotice — inline note for panes whose output is a request header
// (Precondition, Bypass). On PPTB 1.2.7+ the header is sent once the user
// approves it in the host, so this is an info note about that prompt. On
// older hosts it's the warning that pairs with the Execute blocker from
// engine/hostHeaderAdvisories.tsx (#45); `children` carries that copy.

import { MessageBar, MessageBarBody, MessageBarTitle } from '@fluentui/react-components';
import type { ReactNode } from 'react';
import { useHostSession } from '../host/HostContext';
import { hostSupportsAdditionalHeaders } from '../host/pptbBridge';

export function HostHeaderNotice({
  header,
  children,
}: {
  /** The header(s) this pane emits, for the pass-through note. */
  header: ReactNode;
  /** Copy for hosts that can't send the header. */
  children: ReactNode;
}) {
  const host = useHostSession();
  if (!host.embedded) return null;
  if (hostSupportsAdditionalHeaders()) {
    return (
      <MessageBar layout="multiline" intent="info" style={{ marginBottom: 14 }}>
        <MessageBarBody>
          <MessageBarTitle>Sent by Power Platform ToolBox after you approve it</MessageBarTitle>
          Power Platform ToolBox sends {header} with the request. If you haven&apos;t approved
          additional request headers for this tool yet, it asks you first. If you reject, the
          request isn&apos;t sent. You can revoke an approval in ToolBox&apos;s Consent Review.
        </MessageBarBody>
      </MessageBar>
    );
  }
  return (
    <MessageBar layout="multiline" intent="warning" style={{ marginBottom: 14 }}>
      <MessageBarBody>
        <MessageBarTitle>Not sent by this version of Power Platform ToolBox</MessageBarTitle>
        {children}
      </MessageBarBody>
    </MessageBar>
  );
}
