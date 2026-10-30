// ScopeChip — the entity scope, one button in the header after the request
// type (#42, #275; variant A, shared with FetchXML Studio and Command Bar
// Studio). It always shows the scope you're working in and opens a popover
// with the entity source, publishers, solutions and Advanced Find only.
// Changes apply right away, through the same per-org settings the Settings
// drawer used to edit (#52). On first run without a scope, a teaching
// popover points at it, once per environment.

import { useCallback, useEffect, useMemo, useRef, useState, type ReactElement } from 'react';
import {
  Button,
  Caption1,
  Field,
  Popover,
  PopoverSurface,
  PopoverTrigger,
  Radio,
  RadioGroup,
  Switch,
  Tag,
  TagPicker,
  TagPickerControl,
  TagPickerGroup,
  TagPickerInput,
  TagPickerList,
  TagPickerOption,
  TagPickerOptionGroup,
  TeachingPopover,
  TeachingPopoverBody,
  TeachingPopoverFooter,
  TeachingPopoverHeader,
  TeachingPopoverSurface,
  TeachingPopoverTitle,
  Text,
  Tooltip,
  makeStyles,
  tokens,
} from '@fluentui/react-components';
import { ChevronDown16Regular, Filter20Regular, LockClosed16Regular } from '@fluentui/react-icons';
import type { AccessSummary, Publisher, Solution } from '../host/pptbClient';
import { usePersistedSettings, useSettingsLoaded } from '../host/usePersistedSettings';
import { usePublisherFilter } from '../host/usePublisherFilter';
import { useSolutionFilter } from '../host/useSolutionFilter';
import { useActiveOrg } from '../host/activeOrg';
import { useRegisterOpenScope } from '../host/useOpenScope';
import type { EntityScopeMode } from '../state/displaySettings';
import {
  SCOPE_TIP_BODY,
  SCOPE_TIP_TITLE,
  scopeChipText,
  solutionDetail,
  solutionMatchesSearch,
  solutionTagLabel,
} from '../state/scopeChip';

const useStyles = makeStyles({
  chip: {
    flex: '0 1 auto',
    minWidth: 0,
    maxWidth: '420px',
    justifyContent: 'flex-start',
  },
  chipLabel: {
    flex: 1,
    minWidth: 0,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
    textAlign: 'start',
  },
  chevron: { flexShrink: 0, marginLeft: tokens.spacingHorizontalXS },
  tooltipDetail: { whiteSpace: 'pre-line' },
  surface: {
    width: '380px',
    maxWidth: 'calc(100vw - 32px)',
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
  },
  muted: { color: tokens.colorNeutralForeground3 },
  locked: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS,
    color: tokens.colorNeutralForeground3,
  },
  empty: {
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    color: tokens.colorNeutralForeground3,
  },
  // Name, then the detail as a true second line. Fluent's `secondaryContent`
  // renders a squeezed right-hand column instead (#40).
  option: { display: 'flex', flexDirection: 'column', minWidth: 0 },
  optionDetail: { color: tokens.colorNeutralForeground3, wordBreak: 'break-all' },
  tag: { maxWidth: '100%' },
  footer: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: tokens.spacingHorizontalM,
  },
});

export function ScopeChip({ accessSummary }: { accessSummary?: AccessSummary | null }) {
  const styles = useStyles();
  const [settings, updateSettings] = usePersistedSettings();
  const loaded = useSettingsLoaded();
  const org = useActiveOrg();
  const mode = settings.entityScopeMode;
  const { selectedPublisherIds, selectedSolutionIds } = settings;

  // Only the list this mode needs is fetched; both are cached per org.
  const { publishersWithSolutions, publishersLoading, publishersError } = usePublisherFilter(
    mode === 'publisher-solution',
  );
  const {
    solutions: allSolutions,
    solutionsLoading,
    solutionsError,
  } = useSolutionFilter(mode === 'solution-only');

  const publishers = useMemo(
    () => publishersWithSolutions.map((p) => p.publisher),
    [publishersWithSolutions],
  );
  // The solutions to choose from: the selected publishers' in Publisher →
  // solution, every solution with entities in Solution only.
  const solutions = useMemo(
    () =>
      mode === 'publisher-solution'
        ? publishersWithSolutions
            .filter((p) => selectedPublisherIds.includes(p.publisher.publisherid))
            .flatMap((p) => p.solutions)
        : mode === 'solution-only'
          ? allSolutions
          : [],
    [mode, publishersWithSolutions, selectedPublisherIds, allSolutions],
  );
  const selectedPublishers = useMemo(
    () => publishers.filter((p) => selectedPublisherIds.includes(p.publisherid)),
    [publishers, selectedPublisherIds],
  );
  const selectedSolutions = useMemo(
    () => solutions.filter((s) => selectedSolutionIds.includes(s.solutionid)),
    [solutions, selectedSolutionIds],
  );
  const text = scopeChipText(
    mode,
    selectedPublishers,
    selectedSolutions,
    selectedSolutionIds.length,
  );

  const [open, setOpen] = useState(false);
  const [chip, setChip] = useState<HTMLElement | null>(null);
  useRegisterOpenScope(useCallback(() => setOpen(true), []));

  // The full scope as a tooltip, only when the label is cut off.
  const labelRef = useRef<HTMLSpanElement>(null);
  const [truncated, setTruncated] = useState(false);
  const [tooltipVisible, setTooltipVisible] = useState(false);
  useEffect(() => {
    const label = labelRef.current;
    if (!label || typeof ResizeObserver === 'undefined') return;
    const check = () => setTruncated(label.scrollWidth > label.clientWidth);
    check();
    const observer = new ResizeObserver(check);
    observer.observe(label);
    return () => observer.disconnect();
  }, [text.label]);

  // First run: once per environment, until a scope is set. Not before the
  // remembered scope has loaded, so it doesn't flash.
  const dismissedList = Array.isArray(settings.scopeTipDismissed) ? settings.scopeTipDismissed : [];
  const showTip = loaded && !!org && !text.configured && !dismissedList.includes(org);
  const dismissTip = () => {
    if (!org || dismissedList.includes(org)) return;
    updateSettings({ ...settings, scopeTipDismissed: [...dismissedList, org] });
  };

  // Privileges — before the AccessSummary loads, every option is allowed
  // (deferred validation, as the Settings drawer did before #42).
  const canUsePublisherSolution = !accessSummary || accessSummary.fullFilterMode;
  const canUseSolutionOnly =
    !accessSummary || accessSummary.fullFilterMode || accessSummary.solutionsOnlyMode;

  // Removing a publisher also removes its solutions (and only its: adding
  // one leaves the selection alone).
  const onPublishersChange = (ids: string[]) => {
    const dropped = new Set(
      publishersWithSolutions
        .filter((p) => !ids.includes(p.publisher.publisherid))
        .filter((p) => selectedPublisherIds.includes(p.publisher.publisherid))
        .flatMap((p) => p.solutions.map((s) => s.solutionid)),
    );
    updateSettings({
      ...settings,
      selectedPublisherIds: ids,
      selectedSolutionIds: selectedSolutionIds.filter((id) => !dropped.has(id)),
    });
  };

  // Solutions picked in Solution only stay visible in Publisher → solution:
  // their publishers are selected with them.
  const onModeChange = (next: EntityScopeMode) => {
    let publisherIds = selectedPublisherIds;
    if (next === 'publisher-solution' && mode === 'solution-only') {
      const owners = allSolutions
        .filter((s) => selectedSolutionIds.includes(s.solutionid) && s._publisherid_value)
        .map((s) => s._publisherid_value as string);
      publisherIds = [...new Set([...selectedPublisherIds, ...owners])];
    }
    updateSettings({ ...settings, entityScopeMode: next, selectedPublisherIds: publisherIds });
  };

  return (
    <>
      <Popover
        open={open}
        onOpenChange={(_e, d) => setOpen(d.open)}
        positioning="below-start"
        trapFocus
        withArrow
      >
        <PopoverTrigger disableButtonEnhancement>
          <Tooltip
            content={<span className={styles.tooltipDetail}>{text.detail}</span>}
            relationship="description"
            visible={truncated && tooltipVisible && !open}
            onVisibleChange={(_e, d) => setTooltipVisible(d.visible)}
          >
            <Button
              ref={setChip}
              className={styles.chip}
              // Primary only once the remembered scope has loaded and there
              // isn't one, so a returning user doesn't see it flash.
              appearance={text.configured || !loaded ? 'outline' : 'primary'}
              size="small"
              icon={<Filter20Regular />}
              aria-label={text.ariaLabel}
              data-testid="scope-chip"
            >
              <span ref={labelRef} className={styles.chipLabel}>
                {text.label}
              </span>
              <ChevronDown16Regular className={styles.chevron} />
            </Button>
          </Tooltip>
        </PopoverTrigger>
        <PopoverSurface className={styles.surface} aria-label="Entity scope">
          <Text weight="semibold" size={400}>
            Entity scope
          </Text>
          <ScopeEditor
            mode={mode}
            onModeChange={onModeChange}
            canUsePublisherSolution={canUsePublisherSolution}
            canUseSolutionOnly={canUseSolutionOnly}
            publishers={publishers}
            selectedPublishers={selectedPublishers}
            selectedPublisherIds={selectedPublisherIds}
            publishersLoading={publishersLoading}
            publishersError={publishersError}
            onPublishersChange={onPublishersChange}
            solutions={solutions}
            selectedSolutions={selectedSolutions}
            selectedSolutionIds={selectedSolutionIds}
            solutionsLoading={mode === 'solution-only' ? solutionsLoading : publishersLoading}
            solutionsError={mode === 'solution-only' ? solutionsError : null}
            onSolutionsChange={(ids) => updateSettings({ ...settings, selectedSolutionIds: ids })}
            advancedFindOnly={settings.advancedFindOnly}
            onAdvancedFindOnlyChange={(v) => updateSettings({ ...settings, advancedFindOnly: v })}
          />
          <div className={styles.footer}>
            <Caption1 className={styles.muted}>Changes apply right away.</Caption1>
            <Button
              appearance="primary"
              onClick={() => {
                setOpen(false);
                // Back to the chip, as Escape does (a programmatic close
                // doesn't restore focus).
                requestAnimationFrame(() => chip?.focus());
              }}
            >
              Done
            </Button>
          </div>
        </PopoverSurface>
      </Popover>

      {chip && (
        <TeachingPopover
          open={showTip && !open}
          onOpenChange={(_e, d) => {
            if (!d.open) dismissTip();
          }}
          positioning={{ target: chip, position: 'below', align: 'start' }}
        >
          <TeachingPopoverSurface style={{ maxWidth: 320 }}>
            <TeachingPopoverHeader>Tip</TeachingPopoverHeader>
            <TeachingPopoverBody>
              <TeachingPopoverTitle>{SCOPE_TIP_TITLE}</TeachingPopoverTitle>
              <div>{SCOPE_TIP_BODY}</div>
            </TeachingPopoverBody>
            <TeachingPopoverFooter
              primary={{
                children: 'Choose scope',
                onClick: () => {
                  dismissTip();
                  setOpen(true);
                },
              }}
              secondary={{ children: 'Got it', onClick: dismissTip }}
            />
          </TeachingPopoverSurface>
        </TeachingPopover>
      )}
    </>
  );
}

/**
 * Why a scope option is disabled. Only rendered with a hint: an empty Tooltip
 * still shows an empty bubble beside the focused radio.
 */
function PrivilegeHint({ hint, children }: { hint: string; children: ReactElement }) {
  if (!hint) return children;
  return (
    <Tooltip content={hint} relationship="description" positioning="before">
      {children}
    </Tooltip>
  );
}

interface ScopeEditorProps {
  mode: EntityScopeMode;
  onModeChange: (mode: EntityScopeMode) => void;
  canUsePublisherSolution: boolean;
  canUseSolutionOnly: boolean;
  publishers: Publisher[];
  selectedPublishers: Publisher[];
  selectedPublisherIds: string[];
  publishersLoading: boolean;
  publishersError: string | null;
  onPublishersChange: (ids: string[]) => void;
  solutions: Solution[];
  selectedSolutions: Solution[];
  selectedSolutionIds: string[];
  solutionsLoading: boolean;
  solutionsError: string | null;
  onSolutionsChange: (ids: string[]) => void;
  advancedFindOnly: boolean;
  onAdvancedFindOnlyChange: (value: boolean) => void;
}

/** The popover's body: entity source, publishers, solutions, Advanced Find only. */
function ScopeEditor({
  mode,
  onModeChange,
  canUsePublisherSolution,
  canUseSolutionOnly,
  publishers,
  selectedPublishers,
  selectedPublisherIds,
  publishersLoading,
  publishersError,
  onPublishersChange,
  solutions,
  selectedSolutions,
  selectedSolutionIds,
  solutionsLoading,
  solutionsError,
  onSolutionsChange,
  advancedFindOnly,
  onAdvancedFindOnlyChange,
}: ScopeEditorProps) {
  const styles = useStyles();
  const [publisherQuery, setPublisherQuery] = useState('');
  const [solutionQuery, setSolutionQuery] = useState('');

  const publisherOptions = useMemo(() => {
    const q = publisherQuery.trim().toLowerCase();
    return publishers.filter(
      (p) =>
        !selectedPublisherIds.includes(p.publisherid) &&
        (!q ||
          [p.friendlyname, p.uniquename, p.customizationprefix].some((v) =>
            (v ?? '').toLowerCase().includes(q),
          )),
    );
  }, [publishers, selectedPublisherIds, publisherQuery]);

  const { unmanaged, managed } = useMemo(() => {
    const open = solutions.filter(
      (s) => !selectedSolutionIds.includes(s.solutionid) && solutionMatchesSearch(s, solutionQuery),
    );
    return {
      unmanaged: open.filter((s) => !s.ismanaged),
      managed: open.filter((s) => s.ismanaged),
    };
  }, [solutions, selectedSolutionIds, solutionQuery]);

  const needsPublisher = mode === 'publisher-solution' && selectedPublisherIds.length === 0;
  const solutionOption = (s: Solution) => (
    <TagPickerOption key={s.solutionid} value={s.solutionid} text={s.friendlyname}>
      <div className={styles.option}>
        <span>{s.friendlyname}</span>
        <Caption1 className={styles.optionDetail}>{solutionDetail(s)}</Caption1>
      </div>
    </TagPickerOption>
  );

  return (
    <>
      <Field label="Show entities from">
        <RadioGroup value={mode} onChange={(_e, d) => onModeChange(d.value as EntityScopeMode)}>
          <PrivilegeHint
            hint={canUsePublisherSolution ? '' : 'Requires prvReadPublisher and prvReadSolution'}
          >
            <Radio
              value="publisher-solution"
              label="Publisher → solution (default)"
              disabled={!canUsePublisherSolution}
            />
          </PrivilegeHint>
          <PrivilegeHint hint={canUseSolutionOnly ? '' : 'Requires prvReadSolution'}>
            <Radio value="solution-only" label="Solution only" disabled={!canUseSolutionOnly} />
          </PrivilegeHint>
          <Radio value="all" label="All entities" />
        </RadioGroup>
      </Field>

      {mode !== 'all' && (
        <Field
          label="Publishers"
          validationState={publishersError && mode === 'publisher-solution' ? 'error' : 'none'}
          validationMessage={mode === 'publisher-solution' ? publishersError : undefined}
        >
          <TagPicker
            selectedOptions={selectedPublisherIds}
            onOptionSelect={(_e, d) => {
              onPublishersChange(d.selectedOptions);
              setPublisherQuery('');
            }}
            disabled={mode !== 'publisher-solution' || publishersLoading}
          >
            <TagPickerControl>
              <TagPickerGroup aria-label="Selected publishers">
                {mode === 'publisher-solution' &&
                  selectedPublishers.map((p) => (
                    <Tag
                      key={p.publisherid}
                      value={p.publisherid}
                      shape="rounded"
                      size="small"
                      dismissible
                      dismissIcon={{ 'aria-label': 'Remove' }}
                    >
                      {p.friendlyname}
                    </Tag>
                  ))}
              </TagPickerGroup>
              <TagPickerInput
                aria-label="Publishers"
                placeholder={
                  !canUsePublisherSolution
                    ? 'No access to publishers'
                    : publishersLoading
                      ? 'Loading…'
                      : selectedPublisherIds.length
                        ? ''
                        : 'Search publishers…'
                }
                value={publisherQuery}
                onChange={(e) => setPublisherQuery(e.target.value)}
              />
            </TagPickerControl>
            <TagPickerList>
              {publisherOptions.length === 0 ? (
                <div className={styles.empty}>
                  {publisherQuery.trim()
                    ? `No publishers match "${publisherQuery.trim()}"`
                    : 'No publishers found'}
                </div>
              ) : (
                publisherOptions.map((p) => (
                  <TagPickerOption key={p.publisherid} value={p.publisherid} text={p.friendlyname}>
                    <div className={styles.option}>
                      <span>{p.friendlyname}</span>
                      <Caption1 className={styles.optionDetail}>
                        {p.customizationprefix ? `Prefix ${p.customizationprefix}_` : p.uniquename}
                      </Caption1>
                    </div>
                  </TagPickerOption>
                ))
              )}
            </TagPickerList>
          </TagPicker>
          {mode === 'solution-only' && canUsePublisherSolution && (
            <Caption1 className={styles.locked}>
              <LockClosed16Regular /> Not used in Solution only mode
            </Caption1>
          )}
        </Field>
      )}

      {mode !== 'all' && (
        <Field
          label="Solutions"
          validationState={solutionsError ? 'error' : 'none'}
          validationMessage={solutionsError ?? undefined}
        >
          <TagPicker
            selectedOptions={selectedSolutionIds}
            onOptionSelect={(_e, d) => {
              onSolutionsChange(d.selectedOptions);
              setSolutionQuery('');
            }}
            disabled={needsPublisher || solutionsLoading}
          >
            <TagPickerControl>
              <TagPickerGroup aria-label="Selected solutions">
                {selectedSolutions.map((s) => (
                  <Tag
                    key={s.solutionid}
                    value={s.solutionid}
                    shape="rounded"
                    size="small"
                    dismissible
                    dismissIcon={{ 'aria-label': 'Remove' }}
                    title={solutionDetail(s)}
                    className={styles.tag}
                    primaryText={{
                      style: { overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' },
                    }}
                  >
                    {solutionTagLabel(s)}
                  </Tag>
                ))}
              </TagPickerGroup>
              <TagPickerInput
                aria-label="Solutions"
                placeholder={
                  needsPublisher
                    ? 'Select a publisher first'
                    : solutionsLoading
                      ? 'Loading…'
                      : selectedSolutionIds.length
                        ? ''
                        : 'Search name, unique name or version…'
                }
                value={solutionQuery}
                onChange={(e) => setSolutionQuery(e.target.value)}
              />
            </TagPickerControl>
            <TagPickerList>
              {unmanaged.length === 0 && managed.length === 0 ? (
                <div className={styles.empty}>
                  {solutionQuery.trim()
                    ? `No solutions match "${solutionQuery.trim()}"`
                    : 'No solutions found'}
                </div>
              ) : (
                <>
                  {unmanaged.length > 0 && (
                    <TagPickerOptionGroup label="Unmanaged">
                      {unmanaged.map(solutionOption)}
                    </TagPickerOptionGroup>
                  )}
                  {managed.length > 0 && (
                    <TagPickerOptionGroup label="Managed">
                      {managed.map(solutionOption)}
                    </TagPickerOptionGroup>
                  )}
                </>
              )}
            </TagPickerList>
          </TagPicker>
        </Field>
      )}

      <div>
        <Switch
          checked={advancedFindOnly}
          onChange={(_e, d) => onAdvancedFindOnlyChange(d.checked)}
          label="Advanced Find only"
        />
        <Caption1 className={styles.muted} block>
          {advancedFindOnly
            ? "The entity, column and relationship pickers offer only what Advanced Find shows. Requests aren't restricted: a column already in a request keeps working, and Image and File columns are always offered."
            : 'The pickers offer every entity, column and relationship.'}
        </Caption1>
      </div>
    </>
  );
}
