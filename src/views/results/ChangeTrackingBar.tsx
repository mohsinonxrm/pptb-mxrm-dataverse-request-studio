// Change tracking in the results (#233): the response's `@odata.deltaLink`,
// with Get changes (which runs it, through the same path as the next page)
// and Copy delta link.

import {
  Button,
  MessageBar,
  MessageBarActions,
  MessageBarBody,
  MessageBarTitle,
  Spinner,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import { ArrowSync20Regular, Copy20Regular } from '@fluentui/react-icons';
import { copyText } from '../../shell/copy';

const useStyles = makeStyles({
  link: {
    display: 'block',
    marginTop: tokens.spacingVerticalXS,
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground2,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
    minWidth: 0,
  },
});

export function ChangeTrackingBar({
  deltaLink,
  onGetChanges,
  gettingChanges = false,
  isDelta = false,
}: {
  /** The newest `@odata.deltaLink`. */
  deltaLink: string;
  /** Runs the delta link; absent while another request is in flight. */
  onGetChanges?: () => void;
  /** Get changes is running. */
  gettingChanges?: boolean;
  /** The results shown are a delta link's (changes), not the first request's. */
  isDelta?: boolean;
}) {
  const s = useStyles();
  return (
    <MessageBar intent="info" layout="multiline" data-testid="change-tracking-bar">
      <MessageBarBody>
        <MessageBarTitle>
          {isDelta ? 'Changes since the previous delta link' : 'Change tracking is on'}
        </MessageBarTitle>
        {isDelta
          ? 'Get changes again returns the rows created, updated or deleted since these results.'
          : 'Get changes returns the rows created, updated or deleted since these results.'}
        <span className={s.link} title={deltaLink}>
          {deltaLink}
        </span>
      </MessageBarBody>
      <MessageBarActions>
        {/* Default, not primary: Execute stays the surface's one primary action (Fluent 2). */}
        <Button
          icon={gettingChanges ? <Spinner size="tiny" /> : <ArrowSync20Regular />}
          disabled={gettingChanges || !onGetChanges}
          onClick={onGetChanges}
        >
          {gettingChanges ? 'Getting changes…' : 'Get changes'}
        </Button>
        <Button icon={<Copy20Regular />} onClick={() => void copyText(deltaLink, 'the delta link')}>
          Copy delta link
        </Button>
      </MessageBarActions>
    </MessageBar>
  );
}
