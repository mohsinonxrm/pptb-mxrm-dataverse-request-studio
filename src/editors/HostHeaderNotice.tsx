// HostHeaderNotice — inline warning for panes whose output is a request
// header the PPTB host doesn't transmit (Precondition, Bypass). Pairs with
// the Execute blocker from engine/hostHeaderAdvisories.tsx (#45).

import { MessageBar, MessageBarBody, MessageBarTitle } from '@fluentui/react-components';
import type { ReactNode } from 'react';
import { useHostSession } from '../host/HostContext';

export function HostHeaderNotice({ children }: { children: ReactNode }) {
  const host = useHostSession();
  if (!host.embedded) return null;
  return (
    <MessageBar layout="multiline" intent="warning" style={{ marginBottom: 14 }}>
      <MessageBarBody>
        <MessageBarTitle>Not sent by Power Platform ToolBox yet</MessageBarTitle>
        {children}
      </MessageBarBody>
    </MessageBar>
  );
}
