// ODataAnnotationsPanel — the small collapsed section at the bottom of
// a RecordDetailCard that exposes the @odata.* / @Microsoft.Dynamics.CRM.*
// metadata keys WITHOUT polluting the main scalar grid.
//
// Why a separate panel:
//   • The casual reader skips it. The grid above stays clean.
//   • The developer who needs to know the etag, the editlink, or the
//     navigation-property annotation can expand and read.
//
// Closed by default. Single-section Accordion.

import {
  Accordion,
  AccordionItem,
  AccordionHeader,
  AccordionPanel,
  Caption1,
  tokens,
} from '@fluentui/react-components';
import { Info16Regular } from '@fluentui/react-icons';
import type { PartitionedAnnotation } from './detailFieldPartitioner';

export interface ODataAnnotationsPanelProps {
  annotations: PartitionedAnnotation[];
}

export function ODataAnnotationsPanel({ annotations }: ODataAnnotationsPanelProps) {
  if (annotations.length === 0) return null;
  return (
    <Accordion collapsible style={{ marginTop: 16 }}>
      <AccordionItem value="annotations">
        <AccordionHeader expandIconPosition="end" icon={<Info16Regular />}>
          <span style={{ display: 'inline-flex', alignItems: 'center', gap: 8 }}>
            <Caption1 style={{ color: tokens.colorNeutralForeground2, fontWeight: 600 }}>
              OData annotations
            </Caption1>
            <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
              · {annotations.length}
            </Caption1>
          </span>
        </AccordionHeader>
        <AccordionPanel>
          <div
            style={{
              display: 'grid',
              gridTemplateColumns: 'minmax(200px, 320px) 1fr',
              rowGap: 4,
              columnGap: 12,
              fontFamily: tokens.fontFamilyMonospace,
              fontSize: 11,
              color: tokens.colorNeutralForeground2,
            }}
          >
            {annotations.map((a) => (
              <span key={a.key} style={{ display: 'contents' }}>
                <span
                  style={{ color: tokens.colorBrandForeground1, wordBreak: 'break-all' }}
                  title={a.key}
                >
                  {a.key}
                </span>
                <span style={{ wordBreak: 'break-all' }}>
                  {a.value == null ? '—' : String(a.value)}
                </span>
              </span>
            ))}
          </div>
        </AccordionPanel>
      </AccordionItem>
    </Accordion>
  );
}
