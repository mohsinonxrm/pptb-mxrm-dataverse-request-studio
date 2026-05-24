import { type ReactNode, Children, cloneElement, isValidElement, type ReactElement } from 'react';
import { ToggleButton, makeStyles, mergeClasses, tokens } from '@fluentui/react-components';

// A segmented-control composition over v9 ToggleButtons.
//
// Fluent UI v9 doesn't ship a SegmentedControl primitive (per Fluent 2 spec —
// segmented controls are achieved via TabList[subtle] for nav and ToggleButton
// pairs for selection). The v2.2 prototype's preferred "segmented" look —
// a single bordered chassis where two/three ToggleButtons share internal
// hairlines with no gap — is reproduced here without any styling done at the
// caller site.
//
// Usage:
//   <SegmentedToggle ariaLabel="Combinator">
//     <ToggleButton checked={comb === 'and'} onClick={() => setComb('and')}>AND</ToggleButton>
//     <ToggleButton checked={comb === 'or'}  onClick={() => setComb('or')}>OR</ToggleButton>
//   </SegmentedToggle>
//
// The chassis owns the rounded corners + outer border; the buttons inside
// have their border-radius zeroed and a 1px hairline separates them.

const useSegmentedStyles = makeStyles({
  root: {
    display: 'inline-flex',
    flexShrink: 0,
    borderRadius: tokens.borderRadiusMedium,
    overflow: 'hidden',
    borderTopWidth: '1px',
    borderRightWidth: '1px',
    borderBottomWidth: '1px',
    borderLeftWidth: '1px',
    borderTopStyle: 'solid',
    borderRightStyle: 'solid',
    borderBottomStyle: 'solid',
    borderLeftStyle: 'solid',
    borderTopColor: tokens.colorNeutralStroke1,
    borderRightColor: tokens.colorNeutralStroke1,
    borderBottomColor: tokens.colorNeutralStroke1,
    borderLeftColor: tokens.colorNeutralStroke1,
    // Children buttons lose their own border + radius
    '& > button': {
      borderTopWidth: 0,
      borderBottomWidth: 0,
      borderLeftWidth: 0,
      borderRightWidth: 0,
      borderTopLeftRadius: 0,
      borderTopRightRadius: 0,
      borderBottomLeftRadius: 0,
      borderBottomRightRadius: 0,
    },
    // Hairline divider between siblings — emulates a segmented control
    '& > button:not(:first-child)': {
      borderLeftWidth: '1px',
      borderLeftStyle: 'solid',
      borderLeftColor: tokens.colorNeutralStroke1,
    },
  },
});

export function SegmentedToggle({
  children, ariaLabel, className,
}: {
  children: ReactNode;
  ariaLabel: string;
  className?: string;
}) {
  const s = useSegmentedStyles();
  // Coerce children to ToggleButtons so we can ensure consistent size + shape
  const items = Children.toArray(children).map((c) => {
    if (isValidElement(c) && c.type === ToggleButton) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const props = c.props as any;
      return cloneElement(c as ReactElement, {
        size: props.size ?? 'small',
        shape: 'rounded',
      });
    }
    return c;
  });
  return (
    <span role="radiogroup" aria-label={ariaLabel} className={mergeClasses(s.root, className)}>
      {items}
    </span>
  );
}
