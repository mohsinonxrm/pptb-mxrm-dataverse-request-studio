// Declared input arguments of the picked workflow (#4).
//
// Reads the workflow's `xaml` once per selection and lists the InArguments
// it declares (see engine/workflowXaml). Read-only on purpose: the Web API
// ExecuteWorkflow action only has an `EntityId` parameter (the SDK's
// ExecuteWorkflowRequest.InputArguments is documented "for internal use
// only" and isn't exposed as an action parameter), so there is no supported
// body key to send them in. Classic on-demand workflows normally declare
// none beyond the runtime's own InputEntities / CreatedEntities; processes
// with real arguments are Actions, which run from Execute Custom Action.

import { useEffect, useState } from 'react';
import {
  Badge,
  Caption1,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  Spinner,
  tokens,
} from '@fluentui/react-components';
import {
  parseWorkflowInArguments,
  workflowArgumentTypeLabel,
  type WorkflowArgument,
} from '../engine/workflowXaml';

type ArgsState =
  | { status: 'loading' }
  | { status: 'error'; message: string }
  | { status: 'ready'; args: WorkflowArgument[] };

export function WorkflowArgumentsCard({ workflowId }: { workflowId: string }) {
  const [st, setSt] = useState<ArgsState>({ status: 'loading' });

  useEffect(() => {
    if (typeof window === 'undefined' || !window.dataverseAPI?.queryData) {
      setSt({ status: 'ready', args: [] });
      return;
    }
    let cancelled = false;
    setSt({ status: 'loading' });
    window.dataverseAPI
      .queryData(`workflows?$select=xaml&$filter=workflowid eq ${workflowId}`)
      .then((res) => {
        if (cancelled) return;
        const row = (res as { value?: Array<Record<string, unknown>> })?.value?.[0];
        const xaml = typeof row?.xaml === 'string' ? row.xaml : null;
        setSt({ status: 'ready', args: parseWorkflowInArguments(xaml) });
      })
      .catch((e) => {
        if (cancelled) return;
        setSt({ status: 'error', message: e instanceof Error ? e.message : String(e) });
      });
    return () => {
      cancelled = true;
    };
  }, [workflowId]);

  const muted = { color: tokens.colorNeutralForeground3, display: 'block', marginTop: 10 };

  if (st.status === 'loading') {
    return (
      <Caption1 style={{ ...muted, display: 'flex', alignItems: 'center', gap: 6 }}>
        <Spinner size="extra-tiny" /> Reading the workflow&apos;s input arguments…
      </Caption1>
    );
  }
  if (st.status === 'error') {
    return (
      <Caption1 style={muted}>
        Couldn&apos;t read the workflow definition to check for input arguments: {st.message}
      </Caption1>
    );
  }
  if (st.args.length === 0) {
    return (
      <Caption1 style={muted}>
        No input arguments declared — <code>EntityId</code> is the whole request body.
      </Caption1>
    );
  }
  return (
    <MessageBar layout="multiline" intent="warning" style={{ marginTop: 12 }}>
      <MessageBarBody>
        <MessageBarTitle>
          This workflow declares {st.args.length} input argument{st.args.length === 1 ? '' : 's'}
        </MessageBarTitle>
        The Web API <code>ExecuteWorkflow</code> action only takes <code>EntityId</code>, so these
        can&apos;t be set from here; the workflow runs with their default values.
        <ul style={{ margin: '6px 0 0', paddingLeft: 18 }}>
          {st.args.map((a) => (
            <li key={a.name}>
              <code>{a.name}</code>{' '}
              <Badge appearance="outline" size="small">
                {workflowArgumentTypeLabel(a.type)}
              </Badge>
              {a.direction === 'InOut' && (
                <Badge appearance="tint" size="small" style={{ marginLeft: 4 }}>
                  in/out
                </Badge>
              )}
              {a.required && (
                <Badge appearance="tint" color="danger" size="small" style={{ marginLeft: 4 }}>
                  required
                </Badge>
              )}
              {a.description && (
                <span style={{ color: tokens.colorNeutralForeground3 }}> — {a.description}</span>
              )}
            </li>
          ))}
        </ul>
      </MessageBarBody>
    </MessageBar>
  );
}
