import {
  Card, Title3, Caption1, tokens, Button, MessageBar, MessageBarBody, Body1,
} from '@fluentui/react-components';
import { useState } from 'react';
import { useStudioStyles } from '../primitives/styles';
import { Sidebar } from '../shell/Sidebar';
import { MainTabs, type MainTab } from '../shell/MainTabs';
import { UrlBar } from '../shell/UrlBar';
import { ModeShell } from '../shell/ModeShell';
import { findRequestType } from '../registry/requestTypes';
import { groupColorVar } from '../theme/theme';

export function StubMode({ requestId, onPick }: { requestId: string; onPick: (id: string) => void }) {
  const type = findRequestType(requestId);
  const s = useStudioStyles();
  void s;
  const [tab, setTab] = useState<MainTab>('builder');
  return (
    <ModeShell
      sidebar={
        <Sidebar
          type={type}
          urlPreview={`(${type.method.toLowerCase()} preview lands with the ${type.group} pass)`}
          sections={[]}
          activeNode=""
          onSelect={() => undefined}
          recents={[]}
        />
      }
      urlBar={
        <UrlBar
          method={type.method}
          url={`(${type.id}) — preview`}
          executeVerb={type.executeVerb}
          disabledReason={`${type.name} lands in the ${type.group} pass.`}
          loading={false}
          onExecute={() => undefined}
        />
      }
    >
      <MainTabs tab={tab} onTabChange={setTab}>
        <div style={{ display: 'flex', justifyContent: 'center', paddingTop: 40 }}>
          <Card style={{ maxWidth: 640, padding: 32, display: 'flex', flexDirection: 'column', gap: 20 }}>
            <div style={{
              display: 'inline-block', padding: '4px 10px', borderRadius: 999,
              backgroundColor: groupColorVar(type.group), color: '#fff',
              fontSize: 11, fontWeight: 700, letterSpacing: '0.06em',
              textTransform: 'uppercase', alignSelf: 'flex-start',
            }}>
              {type.group} · {type.method}
            </div>
            <Title3>{type.name}</Title3>
            <Body1 style={{ color: tokens.colorNeutralForeground2 }}>{type.sub}</Body1>
            <MessageBar layout="multiline" intent="info">
              <MessageBarBody>
                This mode isn't available yet. Pick another mode from the
                request type switcher above.
              </MessageBarBody>
            </MessageBar>
            <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
              Jump into one of the Read modes below to keep exploring.
            </Caption1>
            <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
              <Button appearance="primary" onClick={() => onPick('retrieve-multiple')}>Retrieve Multiple</Button>
              <Button appearance="outline" onClick={() => onPick('retrieve-single')}>Retrieve Single</Button>
              <Button appearance="outline" onClick={() => onPick('retrieve-nextlink')}>Retrieve NextLink</Button>
              <Button appearance="outline" onClick={() => onPick('predefined-query')}>Predefined Query</Button>
            </div>
          </Card>
        </div>
      </MainTabs>
    </ModeShell>
  );
}
