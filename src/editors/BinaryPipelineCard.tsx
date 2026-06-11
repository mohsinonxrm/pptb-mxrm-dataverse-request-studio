// BinaryPipelineCard — visual sequence of HTTP requests for a binary
// operation. The Builder pane is dominated by the pipeline
// (init → upload-N → commit) with status indicators per step.
//
// Each step shows:
//   ┌──┬─────────┬──────────────────────────────────────┐
//   │ ① │ PATCH   │ Initialize chunked upload           │
//   │   │         │ /accounts(<id>)/sample_contractfile │
//   │   │         │   x-ms-transfer-mode: chunked       │
//   │   │         │ Response 200 · Location: …          │
//   └──┴─────────┴──────────────────────────────────────┘

import {
  Badge,
  Caption1,
  Tooltip,
  Button,
  tokens,
  mergeClasses,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
} from '@fluentui/react-components';
import { Code20Regular, Copy20Regular } from '@fluentui/react-icons';
import { useStudioStyles } from '../primitives/styles';
import { MethodPill } from '../primitives/MethodPill';
import { ENV } from '../mock/environment';
import type { BinaryPipelineStep } from '../engine/binaryBuilders';
import type { HttpMethod } from '../registry/requestTypes';
import type { SizeAdvisory } from '../state/binaryState';

export interface BinaryPipelineCardProps {
  steps: BinaryPipelineStep[];
  /** Title for the card. */
  title?: string;
  /** Inline description. */
  sub?: string;
  /** Pre-flight size advisory — surfaced above the pipeline. */
  sizeAdvisory?: SizeAdvisory | null;
}

export function BinaryPipelineCard({
  steps,
  title = 'Pipeline',
  sub,
  sizeAdvisory,
}: BinaryPipelineCardProps) {
  const s = useStudioStyles();
  return (
    <div className={mergeClasses(s.inlineCard)} style={{ padding: 12, maxWidth: 1040 }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 10 }}>
        <Code20Regular style={{ width: 16, height: 16, color: tokens.colorBrandForeground1 }} />
        <strong style={{ fontSize: 12 }}>{title}</strong>
        <Badge appearance="ghost">
          {steps.length} request{steps.length === 1 ? '' : 's'}
        </Badge>
        {sub && (
          <Caption1 style={{ color: tokens.colorNeutralForeground3, marginLeft: 6 }}>
            {sub}
          </Caption1>
        )}
      </div>

      {sizeAdvisory && (
        <div style={{ marginBottom: 10 }}>
          <MessageBar
            intent={
              sizeAdvisory.level === 'error'
                ? 'error'
                : sizeAdvisory.level === 'warn'
                  ? 'warning'
                  : 'success'
            }
            layout="multiline"
          >
            <MessageBarBody>
              <MessageBarTitle>{sizeAdvisory.headline}</MessageBarTitle>
              {sizeAdvisory.detail}
            </MessageBarBody>
          </MessageBar>
        </div>
      )}

      {steps.length === 0 ? (
        <Caption1
          style={{
            color: tokens.colorNeutralForeground3,
            fontStyle: 'italic',
            display: 'block',
            padding: '8px 4px',
          }}
        >
          No pipeline — pick a target / column / operation.
        </Caption1>
      ) : (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
          {steps.map((step) => (
            <PipelineStepRow key={step.n} step={step} />
          ))}
        </div>
      )}
    </div>
  );
}

function PipelineStepRow({ step }: { step: BinaryPipelineStep }) {
  return (
    <div
      style={{
        padding: 10,
        border: `1px solid ${tokens.colorNeutralStroke2}`,
        borderLeft: `3px solid ${tokens.colorBrandStroke1}`,
        borderRadius: tokens.borderRadiusMedium,
        background: tokens.colorNeutralBackground1,
      }}
    >
      <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
        <span
          style={{
            width: 24,
            height: 24,
            borderRadius: 12,
            background: tokens.colorBrandBackground,
            color: tokens.colorNeutralForegroundOnBrand,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            fontWeight: 700,
            fontSize: 11,
            flexShrink: 0,
          }}
        >
          {step.n}
        </span>
        <MethodPill method={step.method as HttpMethod} size="sm" />
        <div style={{ flex: 1, minWidth: 0 }}>
          <div style={{ fontSize: 12, fontWeight: 600 }}>{step.title}</div>
          <div
            style={{
              fontFamily: tokens.fontFamilyMonospace,
              fontSize: 10,
              color: tokens.colorNeutralForeground3,
              wordBreak: 'break-all',
              marginTop: 2,
            }}
          >
            {step.relativeUrl.startsWith('http') || step.relativeUrl.startsWith('<')
              ? step.relativeUrl
              : `https://${ENV.host}${step.relativeUrl}`}
          </div>
        </div>
        <Tooltip content="Copy URL" relationship="label">
          <Button
            size="small"
            appearance="subtle"
            icon={<Copy20Regular />}
            onClick={() => navigator.clipboard?.writeText(step.relativeUrl)}
          />
        </Tooltip>
      </div>

      {(Object.keys(step.headers).length > 0 || step.body || step.detail) && (
        <div
          style={{
            paddingLeft: 34,
            marginTop: 8,
            display: 'flex',
            flexDirection: 'column',
            gap: 4,
            fontSize: 11,
            color: tokens.colorNeutralForeground3,
          }}
        >
          {Object.keys(step.headers).length > 0 && (
            <div>
              <strong style={{ color: tokens.colorNeutralForeground2 }}>Headers</strong>
              <div
                style={{
                  marginTop: 2,
                  fontFamily: tokens.fontFamilyMonospace,
                  fontSize: 10,
                  lineHeight: 1.55,
                }}
              >
                {Object.entries(step.headers).map(([k, v]) => (
                  <div key={k}>
                    <span style={{ color: tokens.colorBrandForeground2 }}>{k}</span>
                    <span style={{ color: tokens.colorNeutralForeground3 }}>: </span>
                    <span>{v}</span>
                  </div>
                ))}
              </div>
            </div>
          )}
          {step.body && (
            <div>
              <strong style={{ color: tokens.colorNeutralForeground2 }}>Body</strong>
              <div
                style={{
                  marginTop: 2,
                  fontFamily: tokens.fontFamilyMonospace,
                  fontSize: 10,
                  padding: 8,
                  background: tokens.colorNeutralBackground3,
                  borderRadius: tokens.borderRadiusSmall,
                  whiteSpace: 'pre-wrap',
                  wordBreak: 'break-all',
                  maxHeight: 200,
                  overflow: 'auto',
                }}
              >
                {typeof step.body === 'string' ? step.body : JSON.stringify(step.body, null, 2)}
              </div>
            </div>
          )}
          {step.detail && (
            <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>{step.detail}</Caption1>
          )}
        </div>
      )}
    </div>
  );
}
