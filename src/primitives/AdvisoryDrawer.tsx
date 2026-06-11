// AdvisoryDrawer — collapsible badge + popover that aggregates all guardrail
// advisories for the current mode.
//
// Inline MessageBars are reserved for genuinely contextual single-shot
// warnings (e.g. SelectEditor "no $select picked"). Aggregated advisories
// — stripped wildcards, query antipatterns, bypass-header notes,
// flow-bypass notes — flow into this drawer.
//
// Visual grammar:
//   • Lives inside the URL bar footer, between the byte counter and Execute.
//   • Collapsed:  a small ghost-tinted Badge with severity-colored dot
//                 + the highest-severity icon + total count. Tooltip on hover.
//   • Expanded:   a popover positioned above the chip, listing advisories
//                 grouped by severity. Each row is a one-line title + an
//                 optional "Open" button that fires onFocusNode.
//   • Hidden:     when there are 0 advisories the drawer renders nothing.
//
// Errors also surface inline at their source (the editor that detected them)
// AND propagate to <UrlBar disabledReason> via disabledReasonFromAdvisories().
// This drawer is the *aggregate* view, not the *only* view.

import { useState } from 'react';
import {
  Popover,
  PopoverTrigger,
  PopoverSurface,
  Badge,
  Button,
  Tooltip,
  tokens,
  mergeClasses,
} from '@fluentui/react-components';
import {
  Warning20Filled,
  Info20Filled,
  ErrorCircle20Filled,
  Open20Regular,
  ChevronDown16Regular,
} from '@fluentui/react-icons';
import { useStudioStyles } from './styles';
import { bucketAdvisories, type Advisory, type AdvisorySeverity } from './advisories';

export interface AdvisoryDrawerProps {
  advisories: Advisory[];
  /** Mode-supplied focus handler. When user clicks "Open" on an advisory,
   *  the drawer calls onFocusNode(advisory.focusNode) and closes itself. */
  onFocusNode?: (nodeId: string) => void;
}

// Icon + color per severity — kept tiny and consistent with v2.2 chrome.
function severityIcon(s: AdvisorySeverity) {
  switch (s) {
    case 'error':
      return <ErrorCircle20Filled style={{ color: tokens.colorPaletteRedForeground1 }} />;
    case 'warning':
      return <Warning20Filled style={{ color: tokens.colorPaletteDarkOrangeForeground1 }} />;
    case 'info':
      return <Info20Filled style={{ color: tokens.colorBrandForeground1 }} />;
  }
}

function severityBadgeColor(s: AdvisorySeverity): 'danger' | 'warning' | 'informative' {
  return s === 'error' ? 'danger' : s === 'warning' ? 'warning' : 'informative';
}

export function AdvisoryDrawer({ advisories, onFocusNode }: AdvisoryDrawerProps) {
  const s = useStudioStyles();
  const [open, setOpen] = useState(false);
  const { errors, warnings, infos, total } = bucketAdvisories(advisories);

  if (total === 0) return null;

  // The chip's top-level severity is the worst one present.
  const topSeverity: AdvisorySeverity =
    errors.length > 0 ? 'error' : warnings.length > 0 ? 'warning' : 'info';

  // Build the chip tooltip — a one-line summary so users can hover before opening.
  const tipParts: string[] = [];
  if (errors.length) tipParts.push(`${errors.length} blocker${errors.length === 1 ? '' : 's'}`);
  if (warnings.length)
    tipParts.push(`${warnings.length} warning${warnings.length === 1 ? '' : 's'}`);
  if (infos.length) tipParts.push(`${infos.length} info`);
  const tooltip = tipParts.join(' · ');

  return (
    <Popover open={open} onOpenChange={(_, d) => setOpen(d.open)} positioning="above-end">
      <PopoverTrigger disableButtonEnhancement>
        <Tooltip content={`Advisories — ${tooltip}`} relationship="label">
          {/* The chip itself: severity icon + count badge + tiny chevron.
              Sized to sit comfortably in the 44px-tall URL-bar footer. */}
          <button
            type="button"
            className={mergeClasses(s.advisoryChip, open && s.advisoryChipOpen)}
            onClick={() => setOpen((o) => !o)}
            aria-label={`Advisories — ${tooltip}`}
          >
            {severityIcon(topSeverity)}
            <Badge size="small" appearance="tint" color={severityBadgeColor(topSeverity)}>
              {total}
            </Badge>
            <ChevronDown16Regular style={{ color: tokens.colorNeutralForeground3 }} />
          </button>
        </Tooltip>
      </PopoverTrigger>
      <PopoverSurface className={s.advisorySurface}>
        <div className={s.advisoryHeader}>
          <strong style={{ fontSize: 13 }}>Advisories</strong>
          <span style={{ fontSize: 11, color: tokens.colorNeutralForeground3, marginLeft: 'auto' }}>
            {tooltip}
          </span>
        </div>
        {/* Errors first — they're blockers and the user must address them. */}
        {errors.length > 0 && (
          <AdvisorySection
            title="Blockers"
            list={errors}
            onFocusNode={onFocusNode}
            setOpen={setOpen}
          />
        )}
        {warnings.length > 0 && (
          <AdvisorySection
            title="Warnings"
            list={warnings}
            onFocusNode={onFocusNode}
            setOpen={setOpen}
          />
        )}
        {infos.length > 0 && (
          <AdvisorySection title="Notes" list={infos} onFocusNode={onFocusNode} setOpen={setOpen} />
        )}
      </PopoverSurface>
    </Popover>
  );
}

function AdvisorySection({
  title,
  list,
  onFocusNode,
  setOpen,
}: {
  title: string;
  list: Advisory[];
  onFocusNode?: (nodeId: string) => void;
  setOpen: (b: boolean) => void;
}) {
  const s = useStudioStyles();
  return (
    <div className={s.advisorySection}>
      <div className={s.advisorySectionHead}>
        {title} · {list.length}
      </div>
      {list.map((a) => (
        <div key={a.id} className={s.advisoryRow}>
          <span className={s.advisoryRowIcon}>{severityIcon(a.severity)}</span>
          <div style={{ flex: 1, minWidth: 0 }}>
            <div className={s.advisoryTitle}>{a.title}</div>
            {a.body && <div className={s.advisoryBody}>{a.body}</div>}
            {a.learnMoreUrl && (
              <a
                href={a.learnMoreUrl}
                target="_blank"
                rel="noopener noreferrer"
                className={s.advisoryLink}
              >
                Learn more ↗
              </a>
            )}
          </div>
          {a.focusNode && onFocusNode && (
            <Tooltip content="Open the affected pane" relationship="label">
              <Button
                appearance="subtle"
                size="small"
                icon={<Open20Regular />}
                onClick={() => {
                  onFocusNode(a.focusNode!);
                  setOpen(false);
                }}
              />
            </Tooltip>
          )}
        </div>
      ))}
    </div>
  );
}
