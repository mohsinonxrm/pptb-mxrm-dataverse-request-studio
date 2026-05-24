import { tokens } from '@fluentui/react-components';

/**
 * Tokenized rendering for the URL bar — matches the v2.2 spec:
 *   - base path (host + /api/data/v9.2 + entity-set + (id) segment) → muted
 *   - `?` and `&` separators                                         → very muted
 *   - `$key=`                                                        → purple (Fluent v2 "relate" tone)
 *   - value tokens (commas split list values cleanly)               → brand blue
 *   - host portion (https://...)                                    → muted, distinct
 *
 * Renders inline-block segments so the parent can apply ellipsis + scroll.
 */
export function UrlTokens({ url, host }: { url: string; host: string }) {
  // Strip leading host if present
  let rest = url.startsWith(`https://${host}`) ? url.slice(`https://${host}`.length) : url;
  let beforeQ = rest;
  let qs = '';
  const qIdx = rest.indexOf('?');
  if (qIdx >= 0) {
    beforeQ = rest.slice(0, qIdx);
    qs = rest.slice(qIdx + 1);
  }
  const pairs = qs ? qs.split('&') : [];

  const hostColor = tokens.colorNeutralForeground3;
  const baseColor = tokens.colorNeutralForeground2;
  const keyColor = '#8764b8'; // matches the .tk-key purple from v2.2 (Fluent 2 "relate")
  const valColor = tokens.colorBrandForeground1;
  const sepColor = tokens.colorNeutralForeground4;

  return (
    <>
      <span style={{ color: hostColor }}>{`https://${host}`}</span>
      <span style={{ color: baseColor }}>{beforeQ}</span>
      {pairs.length > 0 && <span style={{ color: sepColor }}>?</span>}
      {pairs.map((pair, i) => {
        const eq = pair.indexOf('=');
        const key = eq >= 0 ? pair.slice(0, eq) : pair;
        const value = eq >= 0 ? pair.slice(eq + 1) : '';
        return (
          <span key={i}>
            {i > 0 && <span style={{ color: sepColor }}>&amp;</span>}
            <span style={{ color: keyColor, fontWeight: 600 }}>{key}</span>
            {eq >= 0 && <span style={{ color: sepColor }}>=</span>}
            <span style={{ color: valColor }}>{value}</span>
          </span>
        );
      })}
    </>
  );
}
