// All or nothing (#224) — Associate / Disassociate over a collection-valued
// nav property send one request per target, so a failure part-way leaves
// the earlier targets changed. This option sends them as one change set
// through PPTB's `executeTransaction`: Dataverse applies all of them or
// none. Off by default.

import { Caption1, Switch, tokens } from '@fluentui/react-components';

export interface AllOrNothingOptionProps {
  checked: boolean;
  onChange: (next: boolean) => void;
  /** Targets queued; the option takes effect from two. */
  targets: number;
  /** What each target sends, for the copy. */
  verb: 'POST' | 'DELETE';
}

export function AllOrNothingOption({ checked, onChange, targets, verb }: AllOrNothingOptionProps) {
  const action = verb === 'POST' ? 'linked' : 'unlinked';
  return (
    <div style={{ maxWidth: 720, marginTop: 18 }}>
      <Switch checked={checked} onChange={(_, d) => onChange(!!d.checked)} label="All or nothing" />
      <Caption1
        style={{
          display: 'block',
          marginTop: 2,
          marginLeft: 44,
          color: tokens.colorNeutralForeground3,
        }}
      >
        {checked ? (
          <>
            The {verb} requests go together as one <code>$batch</code> change set: if any target
            fails, Dataverse rolls back the rest, so nothing is {action}.
          </>
        ) : (
          <>
            Off: one {verb} per target, in order. A failure part-way leaves the targets before it{' '}
            {action}. Turn on to send them as one <code>$batch</code> change set instead.
          </>
        )}
        {checked && targets < 2 && ' Takes effect with two or more targets.'}
      </Caption1>
    </div>
  );
}
