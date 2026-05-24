import type { ReactNode } from 'react';
import { TabList, Tab, Badge } from '@fluentui/react-components';
import {
  TextBulletList20Regular, Code20Regular, Table20Regular,
} from '@fluentui/react-icons';
import { useStudioStyles } from '../primitives/styles';

export type MainTab = 'builder' | 'code' | 'results';

export function MainTabs({
  tab, onTabChange, resultCount, children,
}: {
  tab: MainTab;
  onTabChange: (t: MainTab) => void;
  resultCount?: number | null;
  children: ReactNode;
}) {
  const s = useStudioStyles();
  return (
    <main className={s.main}>
      <div className={s.mainTabsBar}>
        <TabList selectedValue={tab} onTabSelect={(_, d) => onTabChange(d.value as MainTab)}>
          <Tab value="builder" icon={<TextBulletList20Regular />}>Builder</Tab>
          <Tab value="code"    icon={<Code20Regular />}>Code</Tab>
          <Tab value="results" icon={<Table20Regular />}>
            <span style={{ display: 'inline-flex', alignItems: 'center', gap: 6 }}>
              Results
              {resultCount != null && <Badge size="extra-small" appearance="tint">{resultCount}</Badge>}
            </span>
          </Tab>
        </TabList>
      </div>
      <div className={s.mainBody}>
        {children}
      </div>
    </main>
  );
}
