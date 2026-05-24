import { makeStyles, tokens } from '@fluentui/react-components';

// One file, everything griffel. Per-primitive class blocks are commented.
// Brand/method colors are read via CSS vars (see index.html), not griffel
// tokens — custom keys don't surface through griffel's `tokens.*` proxy.
export const useStudioStyles = makeStyles({
  // ── Frame ──────────────────────────────────────────────
  frame: {
    height: '100vh',
    width: '100vw',
    maxWidth: '100%',
    display: 'grid',
    gridTemplateRows: 'auto 1fr auto',
    backgroundColor: tokens.colorNeutralBackground2,
    color: tokens.colorNeutralForeground1,
    overflow: 'hidden',
    minWidth: 0,
  },

  // ── FrameHeader (top) — Fluent v9 shadow8 per v2.2 anatomy "Frame shell" rule
  header: {
    display: 'flex',
    alignItems: 'center',
    columnGap: '10px',
    paddingLeft: '16px',
    paddingRight: '12px',
    height: '48px',
    backgroundColor: tokens.colorNeutralBackground1,
    borderBottomWidth: '1px',
    borderBottomStyle: 'solid',
    borderBottomColor: tokens.colorNeutralStroke2,
    flexShrink: 0,
    boxShadow: tokens.shadow8,
    zIndex: 2, // sits above body scroll content for shadow visibility
    minWidth: 0,
    overflow: 'hidden',
  },
  brand: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '8px',
    fontSize: '14px',
    fontWeight: 600,
    color: tokens.colorNeutralForeground1,
    letterSpacing: '-0.01em',
  },
  brandDot: {
    display: 'inline-block',
    width: '10px',
    height: '10px',
    borderRadius: tokens.borderRadiusCircular,
    background: 'linear-gradient(135deg, #0f6cbd 0%, #2899f5 100%)',
  },
  spacer: { flexGrow: 1 },

  // ── Body grid (sidebar + main) ─────────────────────────
  body: {
    display: 'grid',
    gridTemplateColumns: '300px minmax(0, 1fr)',
    minHeight: 0,
    minWidth: 0,
    overflow: 'hidden',
  },
  sidebar: {
    backgroundColor: tokens.colorNeutralBackground1,
    borderRightWidth: '1px',
    borderRightStyle: 'solid',
    borderRightColor: tokens.colorNeutralStroke2,
    display: 'flex',
    flexDirection: 'column',
    overflow: 'hidden',
  },
  sidebarScroll: {
    flexGrow: 1,
    overflowY: 'auto',
    paddingTop: '4px',
    paddingBottom: '12px',
  },

  // ── Mode card (sidebar header, group-color stripe) — shadow2 per v2.2 anatomy
  modeCard: {
    padding: '12px 14px',
    borderLeftWidth: '3px',
    borderLeftStyle: 'solid',
    borderBottomWidth: '1px',
    borderBottomStyle: 'solid',
    borderBottomColor: tokens.colorNeutralStroke2,
    backgroundColor: tokens.colorNeutralBackground2,
    boxShadow: tokens.shadow2,
  },
  modeCardType: {
    fontSize: '10px',
    fontWeight: 700,
    letterSpacing: '0.06em',
    textTransform: 'uppercase',
  },
  modeCardName: {
    fontSize: '14px',
    fontWeight: 600,
    color: tokens.colorNeutralForeground1,
    marginTop: '2px',
  },
  modeCardUrl: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: '10.5px',
    color: tokens.colorNeutralForeground3,
    marginTop: '4px',
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
  },

  // ── Section header inside sidebar ──────────────────────
  // Label uses foreground2 (WCAG AA on dark theme); meta (count text) uses
  // foreground3 for hierarchy.
  sectionH: {
    padding: '14px 14px 4px',
    fontSize: '10px',
    fontWeight: 600,
    letterSpacing: '0.06em',
    textTransform: 'uppercase',
    color: tokens.colorNeutralForeground2,
    display: 'flex',
    alignItems: 'center',
    columnGap: '6px',
  },

  // ── ClauseTreeItem (custom; replaces TreeItemLayout) ───
  ctiList: {
    display: 'flex',
    flexDirection: 'column',
    margin: 0,
    padding: 0,
    listStyle: 'none',
  },
  cti: {
    display: 'grid',
    gridTemplateColumns: '20px 18px 1fr auto',
    columnGap: '8px',
    alignItems: 'center',
    paddingLeft: '14px',
    paddingRight: '12px',
    paddingTop: '4px',
    paddingBottom: '4px',
    minHeight: '28px',
    cursor: 'pointer',
    fontSize: '13px',
    color: tokens.colorNeutralForeground1,
    border: 'none',
    background: 'transparent',
    width: '100%',
    textAlign: 'left',
    fontFamily: 'inherit',
    position: 'relative',
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground1Hover,
    },
  },
  ctiSelected: {
    backgroundColor: tokens.colorBrandBackground2,
    fontWeight: 600,
  },
  ctiSelectedRule: {
    position: 'absolute',
    left: 0,
    top: 0,
    bottom: 0,
    width: '2px',
  },
  ctiDirty: {
    color: tokens.colorPaletteRedForeground1,
    fontWeight: 700,
    fontSize: '13px',
    lineHeight: 1,
    width: '6px',
    textAlign: 'center',
  },
  ctiChev: {
    width: '14px',
    height: '14px',
    color: tokens.colorNeutralForeground3,
    flexShrink: 0,
  },
  ctiIcon: {
    width: '16px',
    height: '16px',
    flexShrink: 0,
  },
  ctiLabel: {
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  ctiCodeLabel: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: '12px',
  },

  // Indentation per nest depth (used by recursive trees like $expand)
  ctiIndent1: { paddingLeft: '32px' },
  ctiIndent2: { paddingLeft: '50px' },
  ctiIndent3: { paddingLeft: '68px' },

  // ── Main pane ──────────────────────────────────────────
  main: {
    display: 'flex',
    flexDirection: 'column',
    minWidth: 0,
    minHeight: 0,
    overflow: 'hidden',
  },
  mainTabsBar: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    paddingLeft: '14px',
    paddingRight: '14px',
    borderBottomWidth: '1px',
    borderBottomStyle: 'solid',
    borderBottomColor: tokens.colorNeutralStroke2,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  mainBody: {
    flexGrow: 1,
    overflowY: 'auto',
    overflowX: 'auto',
    backgroundColor: tokens.colorNeutralBackground2,
    padding: '20px 24px',
    minWidth: 0,
    minHeight: 0,
  },
  paneHead: {
    display: 'flex',
    alignItems: 'center',
    gap: '10px',
    paddingBottom: '14px',
    borderBottomWidth: '1px',
    borderBottomStyle: 'solid',
    borderBottomColor: tokens.colorNeutralStroke2,
    marginBottom: '16px',
    flexWrap: 'wrap',
  },
  paneTitle: {
    fontSize: '20px',
    fontWeight: 600,
    color: tokens.colorNeutralForeground1,
  },
  paneSub: {
    fontSize: '12px',
    color: tokens.colorNeutralForeground3,
    // Long subtitles (e.g. ActionParamForm prose) can expand the header
    // row and push the right-side counter off-edge. Let the sub take the
    // full remaining row when it wraps, but cap line-height so it doesn't
    // push the title baseline.
    lineHeight: 1.45,
    minWidth: 0,
    flexShrink: 1,
  },

  // ── URL bar (footer) — Fluent v9 shadow8 per v2.2 anatomy "Frame shell" rule
  // (footer is the opposite edge of the same frame chrome, so it mirrors the
  // header's elevation rather than dropping a shadow inside the body)
  footer: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    paddingLeft: '12px',
    paddingRight: '12px',
    height: '44px',
    backgroundColor: tokens.colorNeutralBackground1,
    borderTopWidth: '1px',
    borderTopStyle: 'solid',
    borderTopColor: tokens.colorNeutralStroke2,
    flexShrink: 0,
    boxShadow: tokens.shadow8,
    zIndex: 2,
    minWidth: 0,
    overflow: 'hidden',
  },
  methodPill: {
    padding: '4px 10px',
    borderRadius: tokens.borderRadiusCircular,
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: '11px',
    fontWeight: 700,
    letterSpacing: '0.04em',
    color: '#fff',
    flexShrink: 0,
    minWidth: '52px',
    textAlign: 'center',
  },
  urlText: {
    display: 'block',
    flexGrow: 1,
    flexShrink: 1,
    flexBasis: '0%',
    minWidth: 0,
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: '12px',
    color: tokens.colorNeutralForeground1,
    backgroundColor: tokens.colorNeutralBackground2,
    padding: '6px 10px',
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
    cursor: 'text',
    userSelect: 'all',
  },

  // ── Advisory drawer chip + popover ────────────────────────────────
  // Lives in the URL bar footer next to Execute. The chip is a compact
  // affordance — icon + count badge + chevron. Surface is a focused
  // popover with grouped advisories. See AdvisoryDrawer.tsx for usage.
  advisoryChip: {
    display: 'inline-flex',
    alignItems: 'center',
    columnGap: '6px',
    height: '28px',
    paddingTop: '0',
    paddingRight: '8px',
    paddingBottom: '0',
    paddingLeft: '8px',
    borderRadius: tokens.borderRadiusMedium,
    // griffel requires longhand border properties when one of the four colors
    // varies between states (we tweak borderTopColor on :hover).
    borderTopWidth: '1px',
    borderRightWidth: '1px',
    borderBottomWidth: '1px',
    borderLeftWidth: '1px',
    borderTopStyle: 'solid',
    borderRightStyle: 'solid',
    borderBottomStyle: 'solid',
    borderLeftStyle: 'solid',
    borderTopColor: tokens.colorNeutralStroke2,
    borderRightColor: tokens.colorNeutralStroke2,
    borderBottomColor: tokens.colorNeutralStroke2,
    borderLeftColor: tokens.colorNeutralStroke2,
    backgroundColor: tokens.colorNeutralBackground1,
    cursor: 'pointer',
    ':hover': {
      backgroundColor: tokens.colorNeutralBackground1Hover,
      borderTopColor: tokens.colorNeutralStroke1,
      borderRightColor: tokens.colorNeutralStroke1,
      borderBottomColor: tokens.colorNeutralStroke1,
      borderLeftColor: tokens.colorNeutralStroke1,
    },
    ':focus-visible': {
      outlineWidth: '2px',
      outlineStyle: 'solid',
      outlineColor: tokens.colorStrokeFocus2,
      outlineOffset: '1px',
    },
  },
  advisoryChipOpen: {
    backgroundColor: tokens.colorNeutralBackground1Pressed,
    borderTopColor: tokens.colorNeutralStroke1,
    borderRightColor: tokens.colorNeutralStroke1,
    borderBottomColor: tokens.colorNeutralStroke1,
    borderLeftColor: tokens.colorNeutralStroke1,
  },
  advisorySurface: {
    padding: 0,
    maxWidth: '420px',
    minWidth: '320px',
    overflow: 'hidden',
  },
  advisoryHeader: {
    display: 'flex',
    alignItems: 'center',
    padding: '10px 12px',
    borderBottomWidth: '1px',
    borderBottomStyle: 'solid',
    borderBottomColor: tokens.colorNeutralStroke2,
    background: tokens.colorNeutralBackground2,
  },
  advisorySection: {
    paddingBottom: '4px',
    ':not(:last-child)': {
      borderBottomWidth: '1px',
      borderBottomStyle: 'solid',
      borderBottomColor: tokens.colorNeutralStroke3,
    },
  },
  advisorySectionHead: {
    fontSize: '10px',
    fontWeight: 700,
    textTransform: 'uppercase',
    letterSpacing: '0.5px',
    color: tokens.colorNeutralForeground3,
    padding: '8px 12px 4px',
  },
  advisoryRow: {
    display: 'flex',
    alignItems: 'flex-start',
    gap: '8px',
    padding: '6px 12px 8px',
    ':hover': {
      background: tokens.colorNeutralBackground1Hover,
    },
  },
  advisoryRowIcon: {
    flexShrink: 0,
    paddingTop: '2px',
  },
  advisoryTitle: {
    fontSize: '12px',
    fontWeight: 600,
    color: tokens.colorNeutralForeground1,
    lineHeight: 1.4,
  },
  advisoryBody: {
    fontSize: '11px',
    color: tokens.colorNeutralForeground2,
    lineHeight: 1.45,
    marginTop: '2px',
  },
  advisoryLink: {
    fontSize: '11px',
    color: tokens.colorBrandForeground1,
    textDecoration: 'none',
    display: 'inline-block',
    marginTop: '4px',
    ':hover': {
      textDecoration: 'underline',
    },
  },

  // ── Picker chip (request type trigger in header) ───────
  pickerChip: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '8px',
    padding: '4px 10px 4px 8px',
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground3,
    cursor: 'pointer',
    fontSize: '13px',
    fontWeight: 500,
    border: 'none',
    color: tokens.colorNeutralForeground1,
    fontFamily: 'inherit',
    height: '30px',
    minWidth: 0,
    maxWidth: '280px',
    overflow: 'hidden',
    ':hover': { backgroundColor: tokens.colorNeutralBackground3Hover },
  },
  pickerChipLabel: {
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
    minWidth: 0,
  },

  // ── Picker palette (Ctrl+K dialog) ─────────────────────
  paletteSearchWrap: {
    padding: '12px 14px',
    borderBottomWidth: '1px',
    borderBottomStyle: 'solid',
    borderBottomColor: tokens.colorNeutralStroke2,
  },
  paletteList: { maxHeight: '480px', overflowY: 'auto', padding: '6px 0' },
  paletteGroupH: {
    padding: '10px 14px 4px',
    fontSize: '10px',
    fontWeight: 600,
    letterSpacing: '0.06em',
    textTransform: 'uppercase',
    color: tokens.colorNeutralForeground3,
  },
  paletteRow: {
    display: 'grid',
    gridTemplateColumns: '24px 1fr 60px',
    columnGap: '10px',
    alignItems: 'center',
    padding: '8px 14px',
    fontSize: '13px',
    cursor: 'pointer',
    color: tokens.colorNeutralForeground1,
    ':hover': { backgroundColor: tokens.colorNeutralBackground1Hover },
  },
  paletteRowDisabled: {
    opacity: 0.5,
    cursor: 'not-allowed',
  },
  paletteRowMethod: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: '10px',
    fontWeight: 700,
    color: tokens.colorNeutralForeground3,
    textAlign: 'right',
  },
  paletteFoot: {
    padding: '8px 14px',
    borderTopWidth: '1px',
    borderTopStyle: 'solid',
    borderTopColor: tokens.colorNeutralStroke2,
    fontSize: '11px',
    color: tokens.colorNeutralForeground3,
    display: 'flex',
    alignItems: 'center',
    gap: '12px',
    backgroundColor: tokens.colorNeutralBackground2,
  },

  // ── Filter editor ──────────────────────────────────────
  filterPreview: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '8px 12px',
    marginBottom: '12px',
    backgroundColor: tokens.colorNeutralBackground2,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: '11px',
    color: tokens.colorNeutralForeground2,
    overflow: 'hidden',
  },
  filterPreviewKey: {
    fontWeight: 600,
    flexShrink: 0,
  },
  filterPreviewVal: {
    flexGrow: 1,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  filterGroup: {
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderLeftWidth: '3px',
    borderLeftStyle: 'solid',
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground1,
    overflow: 'hidden',
    boxShadow: tokens.shadow2, // resting card elevation per v2.2 anatomy
  },
  filterGroupHeader: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '8px 12px',
    backgroundColor: tokens.colorNeutralBackground2,
    borderBottomWidth: '1px',
    borderBottomStyle: 'solid',
    borderBottomColor: tokens.colorNeutralStroke2,
  },
  filterGroupBody: {
    padding: '8px 10px 4px',
    display: 'flex',
    flexDirection: 'column',
    gap: '6px',
  },
  filterGroupFooter: {
    padding: '4px 10px 10px',
    display: 'flex',
    gap: '6px',
  },
  combToggle: {
    display: 'flex',
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: tokens.borderRadiusMedium,
    overflow: 'hidden',
    height: '24px',
  },
  combBtn: {
    padding: '0 12px',
    height: '24px',
    backgroundColor: 'transparent',
    border: 'none',
    fontFamily: tokens.fontFamilyBase,
    fontSize: '11px',
    fontWeight: 600,
    letterSpacing: '0.04em',
    cursor: 'pointer',
    color: tokens.colorNeutralForeground2,
  },
  combBtnActive: {
    color: '#fff',
  },
  combLabel: {
    fontSize: '11px',
    fontWeight: 600,
    color: tokens.colorNeutralForeground3,
    textTransform: 'uppercase',
    letterSpacing: '0.06em',
  },
  filterRow: {
    display: 'grid',
    // Column 4 (value cell) holds THREE elements inline: optional NotToggle
    // (~50px), optional ValueKindToggle (~50px), and the FilterValueInput
    // itself. The earlier 120px min squeezed the input into the toggles —
    // visible in the user's "Status / equals / [NOT toggle] / value Active"
    // row where the input overlapped the toggle. 220px is the new floor
    // (NotToggle 50 + 6 gap + Input min ~160) — fits the worst case while
    // staying compact when only the input is shown.
    gridTemplateColumns: '20px minmax(150px, 1fr) 160px minmax(220px, 1.2fr) 28px',
    gap: '8px',
    alignItems: 'center',
    padding: '6px 8px',
    backgroundColor: tokens.colorNeutralBackground2,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    // Floor for the row width so the Combobox/operator/value controls don't
    // collapse on top of each other on narrow viewports. Sum of the grid
    // template minimums (20 + 150 + 160 + 220 + 28) + 4 × 8px gap + 16px
    // padding = 626px — round up to give the value input a little breathing
    // room. The pane body has `overflowX: 'auto'` + `minWidth: 0` (see
    // `mainBody`) so a horizontal scrollbar appears when the pane is
    // narrower than this floor instead of mangling the controls.
    minWidth: '660px',
  },
  filterRowNoVal: {
    gridTemplateColumns: '20px minmax(150px, 1fr) 160px 1fr 28px',
    // No-value rows are the same overall width — keep the same floor so
    // adjacent rule cards align visually instead of jaggedly hopping when
    // the user switches an op between scalar and null-check kinds.
    minWidth: '660px',
  },
  filterJoin: {
    fontSize: '10px',
    fontWeight: 700,
    letterSpacing: '0.08em',
    color: tokens.colorNeutralForeground3,
    textTransform: 'uppercase',
    paddingLeft: '8px',
  },
  // Wrapper around the v9 <Divider> used between conditions in a group.
  // The divider visually says "and/or" without yelling — tightens vertical rhythm
  // so adjacent rule cards don't feel disconnected.
  filterJoinDivider: {
    paddingTop: '2px',
    paddingBottom: '2px',
    marginTop: '-2px',
    marginBottom: '-2px',
  },

  // ── NOT toggle (per-condition negation) ────────────────
  notToggle: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '4px',
    padding: '2px 8px',
    height: '24px',
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    backgroundColor: 'transparent',
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: '11px',
    fontWeight: 700,
    color: tokens.colorNeutralForeground3,
    cursor: 'pointer',
    flexShrink: 0,
    ':hover': { backgroundColor: tokens.colorNeutralBackground1Hover },
  },
  notToggleActive: {
    backgroundColor: tokens.colorPaletteRedBackground2,
    color: tokens.colorPaletteRedForeground1,
    borderTopColor: tokens.colorPaletteRedBorder1,
    borderRightColor: tokens.colorPaletteRedBorder1,
    borderBottomColor: tokens.colorPaletteRedBorder1,
    borderLeftColor: tokens.colorPaletteRedBorder1,
  },

  // ── Dataverse function block (.fn from v2.2) ───────────
  fnBlock: {
    display: 'flex',
    flexDirection: 'column',
    gap: '6px',
    padding: '10px 12px',
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground2,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderLeftWidth: '3px',
    borderLeftStyle: 'solid',
    borderLeftColor: '#8764b8', // purple — function family
    // Same floor as filterRow — the fnParamRow inside has a 110px label slot
    // + a Combobox + (sometimes) a second value Input. Without a hard floor
    // those Comboboxes squeeze below their min content on narrow panes.
    minWidth: '660px',
  },
  fnHeader: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    flexWrap: 'wrap',
  },
  fnBadge: {
    fontFamily: tokens.fontFamilyMonospace,
    fontWeight: 700,
    fontSize: '10px',
    padding: '2px 7px',
    borderRadius: tokens.borderRadiusSmall,
    backgroundColor: '#efe8f6',
    color: '#8764b8',
  },
  fnPrefix: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: '11px',
    color: tokens.colorNeutralForeground3,
  },
  fnDescription: {
    fontSize: '11px',
    color: tokens.colorNeutralForeground3,
    fontStyle: 'italic',
  },
  fnParams: {
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
    paddingLeft: '8px',
    paddingTop: '4px',
  },
  fnParamRow: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },
  fnParamLabel: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: '11px',
    color: tokens.colorNeutralForeground3,
    minWidth: '110px',
    flexShrink: 0,
  },

  // ── Lambda block (.lam from v2.2) ──────────────────────
  lamBlock: {
    display: 'flex',
    flexDirection: 'column',
    gap: '6px',
    padding: '10px 12px',
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground2,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderLeftWidth: '3px',
    borderLeftStyle: 'solid',
    borderLeftColor: tokens.colorBrandForeground1,
    // Lambda wraps a nested FilterGroupCard whose rows ARE filterRows — so
    // the floor needs to be at least as wide as filterRow + the lambda's
    // own header (nav-picker + alias chip + any/all toggle). Keep the same
    // 660px floor; the lambda's header is flex-wrap so it doesn't add.
    minWidth: '660px',
  },
  lamHeader: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    flexWrap: 'wrap',
  },
  lamBadge: {
    fontFamily: tokens.fontFamilyMonospace,
    fontWeight: 700,
    fontSize: '13px',
    lineHeight: 1.2,
    padding: '2px 8px',
    borderRadius: tokens.borderRadiusSmall,
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorBrandForeground1,
    minWidth: '22px',
    textAlign: 'center',
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
  },
  lamInner: {
    padding: '6px 10px',
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    marginLeft: '24px',
  },
  lamAliasPrefix: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: '12px',
    color: tokens.colorNeutralForeground3,
    fontWeight: 600,
    paddingRight: '4px',
  },

  // ── Drag handle (visual only) ──────────────────────────
  dragHandle: {
    color: tokens.colorNeutralForeground4,
    width: '14px',
    height: '14px',
    cursor: 'grab',
    flexShrink: 0,
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
  },

  // ── "Generated OData" preview card ─────────────────────
  generatedCard: {
    marginTop: '16px',
    padding: '12px 14px',
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground1,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  generatedLabel: {
    fontSize: '10px',
    fontWeight: 600,
    textTransform: 'uppercase',
    letterSpacing: '0.06em',
    color: tokens.colorNeutralForeground3,
    marginBottom: '6px',
  },
  generatedCode: {
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: '12px',
    color: tokens.colorNeutralForeground1,
    wordBreak: 'break-all',
    lineHeight: 1.55,
  },
  generatedKey: {
    color: '#8764b8',
    fontWeight: 600,
  },

  // ── Add-* menu in the group footer ─────────────────────
  addMenuTrigger: {
    height: '28px',
    padding: '0 10px',
    borderRadius: tokens.borderRadiusMedium,
    border: 'none',
    background: 'transparent',
    fontFamily: 'inherit',
    fontSize: '12px',
    color: tokens.colorBrandForeground1,
    fontWeight: 500,
    cursor: 'pointer',
    display: 'inline-flex',
    alignItems: 'center',
    gap: '4px',
    ':hover': { backgroundColor: tokens.colorNeutralBackground1Hover },
  },

  // ── $select two-list editor ────────────────────────────
  // The two list boxes need bounded height so very-wide tables (e.g. account
  // with 200+ columns) get an INNER vertical scrollbar instead of pushing
  // the whole page taller and forcing the user to scroll the entire pane.
  // We cap to 60vh so the lists always fit on common viewports while leaving
  // room for the header, banner, and PaneHead above. The pane's outer
  // scroll (`mainBody`) still applies if total content overflows.
  selectGrid: {
    display: 'grid',
    gridTemplateColumns: '1fr auto 1fr',
    gap: '12px',
    minHeight: '440px',
    maxHeight: '60vh',
  },
  colList: {
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground1,
    display: 'flex',
    flexDirection: 'column',
    overflow: 'hidden',
    minHeight: 0,
    // Hard cap so the list itself never grows past the grid's maxHeight,
    // and `colListBody` (overflowY: auto) gets a defined containing block
    // to scroll inside.
    maxHeight: '60vh',
  },
  colListH: {
    padding: '8px 12px',
    fontSize: '11px',
    fontWeight: 600,
    color: tokens.colorNeutralForeground2,
    backgroundColor: tokens.colorNeutralBackground2,
    borderBottomWidth: '1px',
    borderBottomStyle: 'solid',
    borderBottomColor: tokens.colorNeutralStroke2,
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
  },
  colListBody: { flexGrow: 1, overflowY: 'auto', padding: '4px 0' },
  colRow: {
    padding: '6px 12px',
    fontSize: '12px',
    cursor: 'pointer',
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    color: tokens.colorNeutralForeground1,
    border: 'none',
    backgroundColor: 'transparent',
    width: '100%',
    textAlign: 'left',
    fontFamily: 'inherit',
    ':hover': { backgroundColor: tokens.colorNeutralBackground1Hover },
  },
  colRowSelected: { backgroundColor: tokens.colorBrandBackground2 },
  colArrows: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    gap: '8px',
  },
  colMeta: {
    fontSize: '10px',
    color: tokens.colorNeutralForeground3,
    fontFamily: tokens.fontFamilyMonospace,
    marginLeft: 'auto',
  },
  colTypeBadge: {
    fontSize: '9px',
    color: tokens.colorNeutralForeground3,
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusSmall,
    padding: '1px 5px',
  },

  // ── Recent runs list ───────────────────────────────────
  recentRow: {
    display: 'flex',
    gap: '8px',
    alignItems: 'center',
    padding: '6px 14px',
    fontSize: '11px',
    color: tokens.colorNeutralForeground2,
    cursor: 'pointer',
    ':hover': { backgroundColor: tokens.colorNeutralBackground1Hover },
  },

  // ── Card row utility ───────────────────────────────────
  inlineCard: {
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground1,
    padding: '14px',
  },

  // Danger variant — red left rule, used for Delete confirmation cards.
  // Per design-review-comprehensive P0-4 / P1-6 (CS-1): consolidates the
  // hand-rolled `borderLeft: 3px solid var(--colorPaletteRedBorderActive)`
  // pattern that appeared across multiple modes.
  inlineCardDanger: {
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderLeftWidth: '3px',
    borderLeftColor: tokens.colorPaletteRedBorderActive,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground1,
    padding: '14px',
  },

  // Accent variant — group-color left rule. Used for Merge Target /
  // Subordinate cards and any other "highlighted but not destructive"
  // card. Caller sets the left-rule color via `style={{ borderLeftColor }}`.
  inlineCardAccent: {
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderLeftWidth: '3px',
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground1,
    padding: '14px',
  },

  // ── Key/value grid ─────────────────────────────────────
  // Per CS-2 — a shared grid for all the "key: value" detail blocks
  // (Workflow detail card, Delete cascade, ManageImage column meta,
  // FunctionUrlPreview, etc.). Two columns, mono key, value flows.
  kvGrid: {
    display: 'grid',
    gridTemplateColumns: '160px 1fr',
    rowGap: '6px',
    columnGap: '14px',
    fontSize: '11px',
    alignItems: 'baseline',
  },
  kvKey: {
    color: tokens.colorNeutralForeground3,
    fontFamily: tokens.fontFamilyBase,
  },
  kvVal: {
    color: tokens.colorNeutralForeground1,
    wordBreak: 'break-word',
  },
  kvValMono: {
    color: tokens.colorNeutralForeground1,
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: '10px',
    wordBreak: 'break-all',
  },

  // ── Detail (Retrieve Single) ───────────────────────────
  detailGrid: {
    display: 'grid',
    gridTemplateColumns: '160px 1fr',
    rowGap: '6px',
    columnGap: '14px',
    alignItems: 'baseline',
  },
  detailKey: {
    fontSize: '11px',
    color: tokens.colorNeutralForeground3,
    textTransform: 'uppercase',
    letterSpacing: '0.04em',
    fontWeight: 600,
  },
  detailVal: {
    fontSize: '13px',
    color: tokens.colorNeutralForeground1,
  },

  // ── Status pill ────────────────────────────────────────
  statusPill: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '4px',
    padding: '3px 8px',
    borderRadius: tokens.borderRadiusCircular,
    fontSize: '11px',
    fontWeight: 600,
    fontFamily: tokens.fontFamilyMonospace,
  },

  // ── Empty state ────────────────────────────────────────
  emptyState: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    padding: '60px 20px',
    textAlign: 'center',
    gap: '12px',
    color: tokens.colorNeutralForeground3,
  },
});
