// Inline amber antipattern indicator for a filter rule or function block.
//
// Discoverability companion to the AdvisoryDrawer in the URL bar: the
// drawer is the canonical "everything DRS wants you to know" view, but
// users need a signal AT THE SOURCE so they know which rule the warning
// pertains to. This icon renders next to the column picker on each rule
// when the column has a known anti-pattern (calculated, logical, large
// text scanned via contains/startswith/endswith).
//
// Hover → Tooltip preview (short).
// Click → Popover with the full title + body + Learn more link to MS docs.
// No antipatterns → returns null (no visual noise on clean rules).

import { Popover, PopoverTrigger, PopoverSurface, Button, Link, tokens } from '@fluentui/react-components';
import { Warning16Filled, Open16Regular } from '@fluentui/react-icons';
import type { ColumnAntipattern } from '../engine/antipatterns';

export interface AntipatternIconProps {
  antipatterns: ColumnAntipattern[];
  /** Optional click-through to focus the URL-bar advisory drawer on this rule.
   *  Currently unused — kept for a future "click → drawer scroll" hook. */
  onFocusDrawer?: () => void;
}

export function AntipatternIcon({ antipatterns }: AntipatternIconProps) {
  if (antipatterns.length === 0) return null;
  return (
    <Popover withArrow positioning="below-end">
      <PopoverTrigger disableButtonEnhancement>
        <Button
          appearance="transparent"
          size="small"
          icon={
            <Warning16Filled
              style={{ color: tokens.colorPaletteDarkOrangeForeground1 }}
            />
          }
          aria-label={`${antipatterns.length} performance warning${antipatterns.length === 1 ? '' : 's'} on this rule`}
          title={
            antipatterns.length === 1
              ? antipatterns[0].title
              : `${antipatterns.length} performance warnings — click for details`
          }
          // Tight sizing so the icon doesn't bump the host row's height.
          // `transparent` appearance + zero padding-block keeps it the
          // same visual height as a 24×24 icon with no button chrome.
          style={{
            minWidth: 20,
            maxWidth: 20,
            width: 20,
            height: 20,
            padding: 0,
            flexShrink: 0,
          }}
        />
      </PopoverTrigger>
      <PopoverSurface style={{ maxWidth: 360, padding: 12 }}>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
          {antipatterns.map((ap, i) => (
            <div key={ap.kind + i} style={{
              display: 'flex',
              flexDirection: 'column',
              gap: 6,
              ...(i > 0 ? {
                paddingTop: 12,
                borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
              } : {}),
            }}>
              <div style={{
                display: 'flex',
                alignItems: 'center',
                gap: 6,
                fontWeight: tokens.fontWeightSemibold,
                fontSize: tokens.fontSizeBase300,
                color: tokens.colorNeutralForeground1,
              }}>
                <Warning16Filled
                  style={{ color: tokens.colorPaletteDarkOrangeForeground1, flexShrink: 0 }}
                />
                <span>{ap.title}</span>
              </div>
              <div style={{
                fontSize: tokens.fontSizeBase200,
                color: tokens.colorNeutralForeground2,
                lineHeight: 1.4,
              }}>
                {ap.body}
              </div>
              <Link
                href={ap.learnMoreUrl}
                target="_blank"
                rel="noopener noreferrer"
                style={{
                  display: 'inline-flex',
                  alignItems: 'center',
                  gap: 4,
                  fontSize: tokens.fontSizeBase200,
                  alignSelf: 'flex-start',
                }}
              >
                Learn more
                <Open16Regular style={{ width: 12, height: 12 }} />
              </Link>
            </div>
          ))}
        </div>
      </PopoverSurface>
    </Popover>
  );
}
