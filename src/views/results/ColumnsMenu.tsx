// The results toolbar's Columns menu (#292): a checkbox per column the
// response has, plus "Show all columns". The choice is saved per table in
// the display settings; the grid and the CSV export follow it.

import {
  Menu,
  MenuButton,
  MenuDivider,
  MenuGroup,
  MenuGroupHeader,
  MenuItem,
  MenuItemCheckbox,
  MenuList,
  MenuPopover,
  MenuTrigger,
  Tooltip,
  makeStyles,
} from '@fluentui/react-components';
import { ColumnTriple20Regular, Eye20Regular } from '@fluentui/react-icons';
import { usePersistedSettings } from '../../host/usePersistedSettings';
import type { ColumnChoice } from './exportResults';
import { columnsButtonLabel, withHiddenColumns } from './hiddenColumns';

const useStyles = makeStyles({
  // Long column lists scroll inside the menu.
  list: { maxHeight: '360px', overflowY: 'auto' },
});

export interface ColumnsMenuProps {
  /** Root table logical name — the choice is saved for this table. */
  table?: string;
  choices: ColumnChoice[];
}

export function ColumnsMenu({ table, choices }: ColumnsMenuProps) {
  const s = useStyles();
  const [settings, updateSettings] = usePersistedSettings();
  const shown = choices.filter((c) => c.shown).map((c) => c.key);
  const allShown = shown.length === choices.length;

  const setHidden = (hidden: string[]) =>
    updateSettings({
      ...settings,
      hiddenColumns: withHiddenColumns(settings.hiddenColumns, table, hidden),
    });

  return (
    <Menu
      persistOnItemClick
      checkedValues={{ columns: shown }}
      onCheckedValueChange={(_, d) => {
        // Keep at least one column.
        if (d.checkedItems.length === 0) return;
        const keep = new Set(d.checkedItems);
        setHidden(choices.filter((c) => !keep.has(c.key)).map((c) => c.key));
      }}
    >
      <MenuTrigger disableButtonEnhancement>
        <Tooltip content="Choose which columns the grid shows" relationship="description">
          <MenuButton
            appearance="subtle"
            size="small"
            icon={<ColumnTriple20Regular />}
            disabled={choices.length === 0}
          >
            {columnsButtonLabel(choices)}
          </MenuButton>
        </Tooltip>
      </MenuTrigger>
      <MenuPopover>
        <MenuList>
          <MenuItem icon={<Eye20Regular />} disabled={allShown} onClick={() => setHidden([])}>
            Show all columns
          </MenuItem>
          <MenuDivider />
          <MenuGroup className={s.list}>
            <MenuGroupHeader>Columns in the response</MenuGroupHeader>
            {choices.map((c) => (
              <MenuItemCheckbox
                key={c.key}
                name="columns"
                value={c.key}
                // The last shown column can't be cleared.
                disabled={c.shown && shown.length === 1}
              >
                {c.label}
              </MenuItemCheckbox>
            ))}
          </MenuGroup>
        </MenuList>
      </MenuPopover>
    </Menu>
  );
}
