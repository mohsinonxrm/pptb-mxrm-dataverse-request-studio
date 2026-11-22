// "Open in Dataverse" — opens a returned record's form in the model-driven
// app (#293). Renders nothing when the record has no GUID, its table has no
// form (an intersect table, an activity without its type, #422) or there's
// no environment.

import { Button, Tooltip, type ButtonProps } from '@fluentui/react-components';
import { Open20Regular } from '@fluentui/react-icons';
import { ENV } from '../env/environment';
import { findTable } from '../metadata/metadata';
import { formTableFor, openInBrowser, recordFormUrl } from '../host/openInDataverse';

export interface OpenInDataverseButtonProps {
  /** Table logical name. */
  table: string | undefined;
  /** The record's GUID. */
  id: unknown;
  /** The record itself: an activity opens as its `activitytypecode`. */
  record?: Record<string, unknown>;
  size?: ButtonProps['size'];
  appearance?: ButtonProps['appearance'];
}

export function OpenInDataverseButton({
  table,
  id,
  record,
  size = 'small',
  appearance = 'subtle',
}: OpenInDataverseButtonProps) {
  const etn = formTableFor(table, record, table ? findTable(table)?.isIntersect : undefined);
  const url = recordFormUrl(ENV.host, etn, id);
  if (!url) return null;
  return (
    <Tooltip
      content="Open this record's form in the model-driven app, in the connection's browser"
      relationship="description"
    >
      <Button
        size={size}
        appearance={appearance}
        icon={<Open20Regular />}
        onClick={() => void openInBrowser(url)}
      >
        Open in Dataverse
      </Button>
    </Tooltip>
  );
}
