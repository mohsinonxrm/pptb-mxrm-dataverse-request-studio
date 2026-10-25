import {
  Fragment,
  useEffect,
  useMemo,
  useRef,
  useState,
  type KeyboardEvent as ReactKeyboardEvent,
} from 'react';
import {
  Popover,
  PopoverTrigger,
  PopoverSurface,
  MenuList,
  MenuItem,
  MenuGroupHeader,
  Dialog,
  DialogSurface,
  DialogBody,
  Input,
  mergeClasses,
  tokens,
  Badge,
} from '@fluentui/react-components';
import { ChevronDown20Regular, Search20Regular } from '@fluentui/react-icons';
import { useStudioStyles } from './styles';
import { groupColorVar } from '../theme/theme';
import { GROUPS, REQ_TYPES, findRequestType, type RequestType } from '../registry/requestTypes';
import { MethodPill } from './MethodPill';

export function RequestTypePicker({
  activeId,
  onPick,
}: {
  activeId: string;
  onPick: (id: string) => void;
}) {
  const s = useStudioStyles();
  const [popOpen, setPopOpen] = useState(false);
  const [palOpen, setPalOpen] = useState(false);
  const [search, setSearch] = useState('');
  const active = findRequestType(activeId);

  // Ctrl+K / ⌘K opens the palette
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'k') {
        e.preventDefault();
        setPalOpen(true);
        setPopOpen(false);
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, []);

  const filtered = useMemo(() => {
    if (!search) return REQ_TYPES;
    const q = search.toLowerCase();
    return REQ_TYPES.filter(
      (r) =>
        r.name.toLowerCase().includes(q) ||
        r.id.includes(q) ||
        r.method.toLowerCase().includes(q) ||
        r.sub.toLowerCase().includes(q),
    );
  }, [search]);

  const grouped = useMemo(() => {
    const m = new Map<string, RequestType[]>(GROUPS.map((g) => [g.id, []]));
    filtered.forEach((r) => m.get(r.group)?.push(r));
    return GROUPS.map((g) => ({ ...g, items: m.get(g.id) ?? [] })).filter((g) => g.items.length);
  }, [filtered]);

  const pick = (id: string) => {
    const t = findRequestType(id);
    if (!t.implemented) {
      // visually allowed; the caller will route to a stub
    }
    onPick(id);
    setPopOpen(false);
    setPalOpen(false);
    setSearch('');
  };

  // Palette keyboard support (#20) — the footer promises ↑↓ / ↵:
  // ArrowDown from the search box enters the list, ArrowUp/Down move between
  // rows (ArrowUp on the first row returns to the search box), Enter in the
  // search box picks the first match, Enter/Space on a row picks it.
  const paletteListRef = useRef<HTMLDivElement>(null);
  const paletteInputRef = useRef<HTMLInputElement>(null);
  const paletteRows = () =>
    Array.from(paletteListRef.current?.querySelectorAll<HTMLElement>('[data-palette-row]') ?? []);
  const onPaletteSearchKeyDown = (e: ReactKeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      paletteRows()[0]?.focus();
    } else if (e.key === 'Enter') {
      const first = grouped[0]?.items[0];
      if (first) {
        e.preventDefault();
        pick(first.id);
      }
    }
  };
  const onPaletteRowKeyDown = (e: ReactKeyboardEvent<HTMLDivElement>, id: string) => {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      pick(id);
      return;
    }
    if (e.key !== 'ArrowDown' && e.key !== 'ArrowUp') return;
    e.preventDefault();
    const rows = paletteRows();
    const i = rows.indexOf(e.currentTarget);
    if (e.key === 'ArrowDown') rows[Math.min(i + 1, rows.length - 1)]?.focus();
    else if (i <= 0) paletteInputRef.current?.focus();
    else rows[i - 1]?.focus();
  };

  const ActiveIcon = active.icon;

  return (
    <Fragment>
      {/* autoSize keeps the request-type list (12 modes) inside the viewport
          on shorter screens — string-form positioning doesn't apply this. */}
      <Popover
        positioning={{ position: 'below', align: 'start', autoSize: 'height-always' }}
        open={popOpen}
        onOpenChange={(_, d) => setPopOpen(d.open)}
        withArrow={false}
      >
        <PopoverTrigger disableButtonEnhancement>
          <button
            className={s.pickerChip}
            type="button"
            aria-haspopup="dialog"
            aria-expanded={popOpen}
          >
            <ActiveIcon
              style={{ color: groupColorVar(active.group), width: 18, height: 18, flexShrink: 0 }}
            />
            <span className={s.pickerChipLabel}>{active.name}</span>
            <MethodPill method={active.method} altMethod={active.altMethod} size="sm" />
            <ChevronDown20Regular
              style={{
                color: tokens.colorNeutralForeground3,
                width: 14,
                height: 14,
                flexShrink: 0,
              }}
            />
          </button>
        </PopoverTrigger>
        <PopoverSurface style={{ padding: 0, minWidth: 460 }}>
          <div
            style={{
              padding: '10px 14px',
              fontSize: 11,
              color: tokens.colorNeutralForeground3,
              display: 'flex',
              justifyContent: 'space-between',
              alignItems: 'center',
              borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
            }}
          >
            <span>Pick request type · {REQ_TYPES.length} modes</span>
            <button
              type="button"
              onClick={() => {
                setPopOpen(false);
                setPalOpen(true);
              }}
              style={{
                display: 'inline-flex',
                alignItems: 'center',
                gap: 4,
                fontSize: 11,
                padding: '2px 8px',
                borderRadius: tokens.borderRadiusMedium,
                background: tokens.colorNeutralBackground3,
                border: 'none',
                color: tokens.colorNeutralForeground2,
                cursor: 'pointer',
                fontFamily: 'inherit',
              }}
            >
              <Search20Regular style={{ width: 12, height: 12 }} /> Search · Ctrl+K
            </button>
          </div>
          <MenuList style={{ padding: 0, maxHeight: 480, overflowY: 'auto' }}>
            {GROUPS.map((g) => {
              const items = REQ_TYPES.filter((r) => r.group === g.id);
              return (
                <Fragment key={g.id}>
                  <MenuGroupHeader>{g.label}</MenuGroupHeader>
                  {items.map((r) => {
                    const Icon = r.icon;
                    return (
                      <MenuItem
                        key={r.id}
                        icon={<Icon style={{ color: groupColorVar(r.group) }} />}
                        onClick={() => pick(r.id)}
                        secondaryContent={
                          <span style={{ display: 'inline-flex', gap: 6, alignItems: 'center' }}>
                            {!r.implemented && (
                              <Badge size="extra-small" appearance="ghost">
                                soon
                              </Badge>
                            )}
                            <span
                              style={{
                                fontFamily: tokens.fontFamilyMonospace,
                                fontSize: 10,
                                fontWeight: 700,
                              }}
                            >
                              {r.method}
                            </span>
                          </span>
                        }
                      >
                        {r.name}
                      </MenuItem>
                    );
                  })}
                </Fragment>
              );
            })}
          </MenuList>
        </PopoverSurface>
      </Popover>

      <Dialog open={palOpen} onOpenChange={(_, d) => setPalOpen(d.open)} modalType="non-modal">
        <DialogSurface
          aria-label="Search request types"
          style={{ padding: 0, maxWidth: 600, width: '90vw' }}
        >
          <DialogBody style={{ display: 'block', padding: 0 }}>
            <div className={s.paletteSearchWrap}>
              <Input
                appearance="filled-darker"
                aria-label="Search request types"
                placeholder="Search request types…"
                value={search}
                onChange={(_, d) => setSearch(d.value)}
                onKeyDown={onPaletteSearchKeyDown}
                input={{ ref: paletteInputRef }}
                contentBefore={<Search20Regular />}
                style={{ width: '100%' }}
                autoFocus
              />
            </div>
            <div className={s.paletteList} ref={paletteListRef}>
              {grouped.length === 0 && (
                <div
                  style={{
                    padding: '20px 14px',
                    color: tokens.colorNeutralForeground3,
                    fontSize: 12,
                  }}
                >
                  No matches.
                </div>
              )}
              {grouped.map((g) => (
                <Fragment key={g.id}>
                  <div className={s.paletteGroupH}>{g.label}</div>
                  {g.items.map((r) => {
                    const Icon = r.icon;
                    return (
                      <div
                        key={r.id}
                        className={mergeClasses(
                          s.paletteRow,
                          !r.implemented && s.paletteRowDisabled,
                        )}
                        onClick={() => pick(r.id)}
                        onKeyDown={(e) => onPaletteRowKeyDown(e, r.id)}
                        role="button"
                        tabIndex={0}
                        data-palette-row
                      >
                        <Icon style={{ color: groupColorVar(r.group), width: 18, height: 18 }} />
                        <div>
                          <div
                            style={{
                              fontWeight: r.id === activeId ? 600 : 400,
                              display: 'flex',
                              alignItems: 'center',
                              gap: 6,
                            }}
                          >
                            {r.name}
                            {!r.implemented && (
                              <Badge size="extra-small" appearance="ghost">
                                soon
                              </Badge>
                            )}
                          </div>
                          <div style={{ fontSize: 11, color: tokens.colorNeutralForeground3 }}>
                            {r.sub}
                          </div>
                        </div>
                        <div className={s.paletteRowMethod}>{r.method}</div>
                      </div>
                    );
                  })}
                </Fragment>
              ))}
            </div>
            <div className={s.paletteFoot}>
              <span>
                <kbd>↑↓</kbd> navigate
              </span>
              <span>
                <kbd>↵</kbd> select
              </span>
              <span>
                <kbd>Esc</kbd> close
              </span>
              <span style={{ marginLeft: 'auto' }}>
                {filtered.length} of {REQ_TYPES.length} types
              </span>
            </div>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </Fragment>
  );
}
