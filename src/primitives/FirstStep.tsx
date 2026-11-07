// The Builder's first-step empty state (#297): what this request type does
// (the request-type picker's description), the first step, and a button that
// opens that step's picker. Each mode shows it until the first step is done.
//
// The mode comes from ModeIdContext, which App's ModeRouter provides, so
// shared editors (TargetEditor, the action pickers) can show it without
// every mode passing its name down.

import { createContext, useContext } from 'react';
import { Body1, Button, Caption1, Subtitle2, makeStyles, tokens } from '@fluentui/react-components';
import { findRequestType } from '../registry/requestTypes';
import { groupColorVar } from '../theme/theme';

/** The routed mode's request-type id ('' outside a mode). */
export const ModeIdContext = createContext('');

const useStyles = makeStyles({
  root: {
    display: 'flex',
    alignItems: 'flex-start',
    gap: tokens.spacingHorizontalL,
    maxWidth: '560px',
    padding: `${tokens.spacingVerticalL} ${tokens.spacingHorizontalL}`,
    marginBottom: tokens.spacingVerticalL,
    borderRadius: tokens.borderRadiusLarge,
    border: `1px dashed ${tokens.colorNeutralStroke1}`,
    backgroundColor: tokens.colorNeutralBackground2,
  },
  icon: { fontSize: '32px', width: '32px', height: '32px', flexShrink: 0 },
  text: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXS, minWidth: 0 },
  sub: { color: tokens.colorNeutralForeground2 },
  step: { color: tokens.colorNeutralForeground3, marginTop: tokens.spacingVerticalXS },
  action: { alignSelf: 'flex-start', marginTop: tokens.spacingVerticalS },
});

export interface FirstStepProps {
  /** The first step, as a button label: "Choose a table", "Choose an action". */
  step: string;
  /** Opens that step's picker. */
  onStart: () => void;
}

export function FirstStep({ step, onStart }: FirstStepProps) {
  const s = useStyles();
  const type = findRequestType(useContext(ModeIdContext));
  const Icon = type.icon;
  return (
    <section className={s.root} aria-label={`Get started with ${type.name}`}>
      <Icon className={s.icon} style={{ color: groupColorVar(type.group) }} aria-hidden />
      <div className={s.text}>
        <Subtitle2>{type.name}</Subtitle2>
        <Body1 className={s.sub}>{type.sub}.</Body1>
        <Caption1 className={s.step}>Step 1 · {step}</Caption1>
        <Button appearance="primary" className={s.action} onClick={onStart}>
          {step}
        </Button>
      </div>
    </section>
  );
}
