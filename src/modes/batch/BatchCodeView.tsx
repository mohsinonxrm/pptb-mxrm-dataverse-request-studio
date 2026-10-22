// Code tab for the $batch composer (#9) — a format picker over plain text.
// The body is CRLF-delimited multipart text, not source a highlighter adds
// much to, so this stays a <pre> (no Monaco chunk for this lazy mode).

import { useState } from 'react';
import { Button, Caption1, Tab, TabList, Tooltip, tokens } from '@fluentui/react-components';
import { Code20Regular, Copy20Regular } from '@fluentui/react-icons';
import { PaneHead } from '../../editors/PaneHead';
import { copyToClipboardSafe } from '../../host/pptbBridge';

export function BatchCodeView<F extends string>({
  formats,
  labels,
  generate,
  note,
  crlfFormats = ['http' as F],
}: {
  formats: F[];
  labels: Record<F, string>;
  generate: (fmt: F) => string;
  note?: string;
  /** Formats that show the raw multipart body — only these get the CRLF note (#198). */
  crlfFormats?: F[];
}) {
  const [fmt, setFmt] = useState<F>(formats[0]);
  const [copied, setCopied] = useState(false);
  const code = generate(fmt);
  return (
    <div>
      <PaneHead icon={Code20Regular} title="Code" sub={note} group="batch">
        <Tooltip content="Copy" relationship="label">
          <Button
            icon={<Copy20Regular />}
            appearance="subtle"
            onClick={() =>
              void copyToClipboardSafe(code).then(() => {
                setCopied(true);
                setTimeout(() => setCopied(false), 1500);
              })
            }
          >
            {copied ? 'Copied' : 'Copy'}
          </Button>
        </Tooltip>
      </PaneHead>
      <TabList selectedValue={fmt} onTabSelect={(_, d) => setFmt(d.value as F)} size="small">
        {formats.map((f) => (
          <Tab key={f} value={f}>
            {labels[f]}
          </Tab>
        ))}
      </TabList>
      <pre
        style={{
          fontFamily: tokens.fontFamilyMonospace,
          fontSize: 12,
          lineHeight: 1.5,
          background: tokens.colorNeutralBackground1,
          border: `1px solid ${tokens.colorNeutralStroke2}`,
          borderRadius: 6,
          padding: 12,
          marginTop: 8,
          whiteSpace: 'pre',
          overflow: 'auto',
          maxHeight: 'calc(100vh - 260px)',
        }}
      >
        {code}
      </pre>
      {crlfFormats.includes(fmt) && (
        <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
          Line breaks in a $batch body are CRLF on the wire.
        </Caption1>
      )}
    </div>
  );
}
