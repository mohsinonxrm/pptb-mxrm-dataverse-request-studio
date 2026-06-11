// RequestBodyPreview — formatted, monospace, copyable JSON body for the
// Builder pane. Every action mode surfaces a "Request body preview" card
// under the parameters so the user sees the wire shape building up as they edit.

import { Caption1, Tooltip, Button, Badge, tokens, mergeClasses } from '@fluentui/react-components';
import { Copy20Regular, Code20Filled } from '@fluentui/react-icons';
import { useStudioStyles } from '../primitives/styles';

export interface RequestBodyPreviewProps {
  /** The JSON body to preview. */
  body: Record<string, unknown>;
  /** Title for the card (defaults to "Request body preview"). */
  title?: string;
  /** Pill text shown next to the title (e.g. "from $metadata"). */
  pillText?: string;
  /** Maximum height before the body scrolls. */
  maxHeight?: number | string;
}

export function RequestBodyPreview({
  body,
  title = 'Request body preview',
  pillText,
  maxHeight = 360,
}: RequestBodyPreviewProps) {
  const s = useStudioStyles();
  const json = JSON.stringify(body, null, 2);
  const isEmpty = Object.keys(body).length === 0;

  return (
    <div className={mergeClasses(s.inlineCard)} style={{ padding: 12, maxWidth: 980 }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 10 }}>
        <Code20Filled style={{ width: 16, height: 16, color: tokens.colorBrandForeground1 }} />
        <strong style={{ fontSize: 12 }}>{title}</strong>
        {pillText && <Badge appearance="ghost">{pillText}</Badge>}
        {isEmpty && (
          <Badge appearance="tint" color="subtle">
            empty
          </Badge>
        )}
        <span style={{ flexGrow: 1 }} />
        <Tooltip content="Copy body JSON" relationship="label">
          <Button
            size="small"
            appearance="subtle"
            icon={<Copy20Regular />}
            onClick={() => navigator.clipboard?.writeText(json)}
            disabled={isEmpty}
          />
        </Tooltip>
      </div>
      {isEmpty ? (
        <Caption1
          style={{
            color: tokens.colorNeutralForeground3,
            fontStyle: 'italic',
            display: 'block',
            padding: '8px 4px',
          }}
        >
          No parameters set yet. The wire body will appear here as you fill in the form above.
        </Caption1>
      ) : (
        <pre
          style={{
            margin: 0,
            padding: 10,
            fontFamily: tokens.fontFamilyMonospace,
            fontSize: 11,
            background: tokens.colorNeutralBackground3,
            borderRadius: tokens.borderRadiusSmall,
            maxHeight,
            overflow: 'auto',
            whiteSpace: 'pre',
          }}
        >
          {json}
        </pre>
      )}
    </div>
  );
}
