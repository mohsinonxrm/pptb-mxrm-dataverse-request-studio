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
      {/* What the request does, for people new to the Web API (#289). The
          card isn't focusable, so a tooltip alone would be mouse-only: the
          sentence is also in the card's text for screen readers (#422), and
          the tooltip, shown on hover, isn't announced a second time. */}
      <Tooltip content={type.description} relationship="inaccessible" positioning="below">
        <div className={s.modeCardName}>
          {type.name}
          <span className={s.srOnly}>. {type.description}</span>
        </div>
      </Tooltip>
      {/* A v9 Tooltip rather than the native title attribute, so the chrome
          matches every other hover surface in the app. Hover only: the URL
          is the card's own text. */}
      <Tooltip content={urlPreview} relationship="description" positioning="below">
        <div className={s.modeCardUrl}>{urlPreview}</div>
      </Tooltip>
    </div>
  );
}
