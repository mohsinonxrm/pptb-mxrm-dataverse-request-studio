// FunctionUrlPreview — colorized URL preview card for the Execute Function
// Builder pane. URL shape:
//
//   <Function>(  <param>=<alias>  ) ? <alias>=<value>
//
// We split the URL into function name, parameter pairs (path part), and
// alias substitutions (query part) so we can color them distinctly. The
// alias form is shown when `useParamAliases` is on; inline form when off.

import { Caption1, Tooltip, Button, tokens, mergeClasses, Badge } from '@fluentui/react-components';
import { Copy20Regular, Link20Filled } from '@fluentui/react-icons';
import { useStudioStyles } from '../primitives/styles';
import type { BuiltRequest } from '../engine/urlBuilder';

export interface FunctionUrlPreviewProps {
  built: BuiltRequest;
  useParamAliases: boolean;
}

export function FunctionUrlPreview({ built, useParamAliases }: FunctionUrlPreviewProps) {
  const s = useStudioStyles();
  // built.relativeUrl looks like:
  //   /api/data/v9.2/<...>/FunctionName(p1=@p1,p2=@p2)?@p1='val'&@p2=42
  // We want to peel out the path + alias substitutions.
  const url = built.relativeNoBase || built.relativeUrl;
  const m = url.match(/^([^?]*)\??(.*)$/);
  const pathPart = m?.[1] ?? url;
  const aliasPart = m?.[2] ?? '';

  // Extract the "FunctionName(...)" segment from pathPart so we can highlight
  // the function name + bind segment separately.
  const fnMatch = pathPart.match(/^(.*?)\/([^/]+)\((.*)\)$/);
  const prefix = fnMatch?.[1] ?? '';
  const fnName = fnMatch?.[2] ?? '';
  const paramSegment = fnMatch?.[3] ?? '';

  const aliasPairs = aliasPart
    ? aliasPart.split('&').map((p) => {
        const eq = p.indexOf('=');
        return eq === -1 ? { key: p, value: '' } : { key: p.slice(0, eq), value: p.slice(eq + 1) };
      })
    : [];

  return (
    <div className={mergeClasses(s.inlineCard)} style={{ padding: 12, maxWidth: 980 }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 10 }}>
        <Link20Filled style={{ width: 16, height: 16, color: tokens.colorBrandForeground1 }} />
        <strong style={{ fontSize: 12 }}>URL preview</strong>
        <Badge appearance="ghost">{useParamAliases ? 'param aliases' : 'inline literals'}</Badge>
        <span style={{ flexGrow: 1 }} />
        <Tooltip content="Copy URL" relationship="label">
          <Button
            size="small"
            appearance="subtle"
            icon={<Copy20Regular />}
            onClick={() => navigator.clipboard?.writeText(built.relativeUrl)}
          />
        </Tooltip>
      </div>

      <div
        style={{
          padding: 10,
          background: tokens.colorNeutralBackground3,
          borderRadius: tokens.borderRadiusSmall,
          fontFamily: tokens.fontFamilyMonospace,
          fontSize: 11,
          wordBreak: 'break-all',
          whiteSpace: 'pre-wrap',
          lineHeight: 1.6,
        }}
      >
        <span style={{ color: tokens.colorNeutralForeground3 }}>{prefix}/</span>
        <span style={{ color: tokens.colorBrandForeground1, fontWeight: 700 }}>
          {fnName || pathPart}
        </span>
        {paramSegment !== '' && (
          <>
            <span style={{ color: tokens.colorNeutralForeground3 }}>(</span>
            <ParamSegment segment={paramSegment} />
            <span style={{ color: tokens.colorNeutralForeground3 }}>)</span>
          </>
        )}
        {aliasPairs.length > 0 && (
          <>
            <span style={{ color: tokens.colorNeutralForeground3 }}>?</span>
            {aliasPairs.map((p, i) => (
              <span key={p.key + i}>
                {i > 0 && <span style={{ color: tokens.colorNeutralForeground3 }}>&</span>}
                <span style={{ color: tokens.colorPaletteDarkOrangeForeground1 }}>{p.key}</span>
                <span style={{ color: tokens.colorNeutralForeground3 }}>=</span>
                <span style={{ color: tokens.colorPaletteGreenForeground1 }}>
                  {decodeURIComponent(p.value)}
                </span>
              </span>
            ))}
          </>
        )}
      </div>

      {useParamAliases && aliasPairs.length > 0 && (
        <Caption1 style={{ display: 'block', marginTop: 8, color: tokens.colorNeutralForeground3 }}>
          Aliased — each{' '}
          <code style={{ color: tokens.colorPaletteDarkOrangeForeground1 }}>
            @p<i>n</i>
          </code>{' '}
          in the path is bound by the query string. Recommended by docs for any non-trivial value.
        </Caption1>
      )}
      {!useParamAliases && (
        <Caption1 style={{ display: 'block', marginTop: 8, color: tokens.colorNeutralForeground3 }}>
          Inline — values are embedded directly in the path. Beware URL-length limits +
          DateTimeOffset encoding bugs.
        </Caption1>
      )}
    </div>
  );
}

// Color each `key=value` pair inside the param segment.
function ParamSegment({ segment }: { segment: string }) {
  const pairs = segment.split(',').map((p) => {
    const eq = p.indexOf('=');
    return eq === -1 ? { key: p, value: '' } : { key: p.slice(0, eq), value: p.slice(eq + 1) };
  });
  return (
    <>
      {pairs.map((p, i) => (
        <span key={p.key + i}>
          {i > 0 && <span style={{ color: tokens.colorNeutralForeground3 }}>,</span>}
          <span style={{ color: tokens.colorNeutralForeground2 }}>{p.key}</span>
          <span style={{ color: tokens.colorNeutralForeground3 }}>=</span>
          <span style={{ color: tokens.colorPaletteDarkOrangeForeground1 }}>{p.value}</span>
        </span>
      ))}
    </>
  );
}
