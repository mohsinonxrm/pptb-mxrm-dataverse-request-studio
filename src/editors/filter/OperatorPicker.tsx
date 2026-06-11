import { useMemo } from 'react';
import {
  Menu,
  MenuTrigger,
  MenuPopover,
  MenuList,
  MenuItem,
  MenuGroupHeader,
  MenuDivider,
  Button,
  tokens,
  Tooltip,
} from '@fluentui/react-components';
import { ChevronDown20Regular } from '@fluentui/react-icons';
import { OP_CATEGORIES, operatorsFor, findOperator, type OpCategory } from './operators';
import type { ColumnMeta } from '../../mock/metadata';

export function OperatorPicker({
  table,
  col,
  value,
  onChange,
  size = 'small',
  only,
}: {
  table: string;
  col: ColumnMeta | undefined;
  value: string;
  onChange: (opId: string) => void;
  size?: 'small' | 'medium';
  /** Restrict to a list of categories (e.g. ['comparison','string'] for plain rule rows) */
  only?: OpCategory[];
}) {
  const ops = useMemo(() => {
    const all = col ? operatorsFor(col.attributeType, table) : [];
    return only ? all.filter((o) => only.includes(o.category)) : all;
  }, [col, table, only]);

  const grouped = useMemo(() => {
    const m = new Map<OpCategory, typeof ops>();
    ops.forEach((o) => {
      if (!m.has(o.category)) m.set(o.category, []);
      m.get(o.category)!.push(o);
    });
    return OP_CATEGORIES.filter((c) => m.has(c.id)).map((c) => ({ ...c, items: m.get(c.id)! }));
  }, [ops]);

  const cur = findOperator(value);

  return (
    // `autoSize: 'height-always'` tells Fluent's Floating-UI positioning to
    // constrain the popover height to the available space between the
    // trigger and the viewport edge. Without this, a tall operator list
    // (especially DateTime which exposes ~30 fns across 6 categories)
    // renders at its natural height, overflows the viewport, and either
    // pushes page layout down or clips. The shortcut string `"below-start"`
    // doesn't apply autoSize; the structured object does.
    <Menu
      hasIcons={false}
      positioning={{ position: 'below', align: 'start', autoSize: 'height-always' }}
    >
      <MenuTrigger disableButtonEnhancement>
        <Button
          size={size}
          appearance="outline"
          iconPosition="after"
          icon={<ChevronDown20Regular />}
          style={{ minWidth: 160, justifyContent: 'space-between' }}
        >
          <span style={{ fontWeight: 500, color: tokens.colorNeutralForeground1 }}>
            {cur?.label ?? value}
          </span>
        </Button>
      </MenuTrigger>
      <MenuPopover style={{ maxHeight: 460, overflowY: 'auto' }}>
        <MenuList>
          {grouped.map((g, i) => (
            <span key={g.id}>
              {i > 0 && <MenuDivider />}
              <MenuGroupHeader>{g.label}</MenuGroupHeader>
              {g.items.map((o) => (
                <Tooltip
                  key={o.id}
                  content={o.hint ?? o.label}
                  relationship="description"
                  positioning="after"
                >
                  <MenuItem onClick={() => onChange(o.id)}>
                    <span
                      style={{
                        display: 'flex',
                        justifyContent: 'space-between',
                        gap: 12,
                        width: '100%',
                      }}
                    >
                      <span style={{ fontWeight: o.id === value ? 600 : 400 }}>{o.label}</span>
                      <span
                        style={{
                          fontFamily: tokens.fontFamilyMonospace,
                          fontSize: 10,
                          color: tokens.colorNeutralForeground3,
                        }}
                      >
                        {o.odata.replace('Microsoft.Dynamics.CRM.', '')}
                      </span>
                    </span>
                  </MenuItem>
                </Tooltip>
              ))}
            </span>
          ))}
        </MenuList>
      </MenuPopover>
    </Menu>
  );
}
