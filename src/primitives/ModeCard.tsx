import { Tooltip } from '@fluentui/react-components';
import { useStudioStyles } from './styles';
import { groupColorVar } from '../theme/theme';
import type { RequestType } from '../registry/requestTypes';

export function ModeCard({ type, urlPreview }: { type: RequestType; urlPreview: string }) {
  const s = useStudioStyles();
  const groupColor = groupColorVar(type.group);
  return (
    <div className={s.modeCard} style={{ borderLeftColor: groupColor }}>
      <div className={s.modeCardType} style={{ color: groupColor }}>
        {type.group} · {type.altMethod ? `${type.method}/${type.altMethod}` : type.method}
      </div>
      <div className={s.modeCardName}>{type.name}</div>
      {/* Use v9 Tooltip instead of the native title attribute so the chrome
          matches every other hover surface in the app, and we get
          keyboard-focus tooltips for free. */}
      <Tooltip content={urlPreview} relationship="description" positioning="below">
        <div className={s.modeCardUrl}>
          {urlPreview}
        </div>
      </Tooltip>
    </div>
  );
}
