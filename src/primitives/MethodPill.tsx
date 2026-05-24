import { useStudioStyles } from './styles';
import { methodColorVar } from '../theme/theme';
import { tokens } from '@fluentui/react-components';
import type { HttpMethod } from '../registry/requestTypes';

export function MethodPill({
  method, altMethod, size = 'md',
}: {
  method: HttpMethod;
  /** Optional secondary verb — when set, the pill renders both with a slash
   *  separator (e.g. "POST/PATCH") for cardinality-dependent modes. */
  altMethod?: HttpMethod;
  size?: 'sm' | 'md';
}) {
  const s = useStudioStyles();
  // Dual-verb pill — render as two-color split so each verb retains its
  // semantic color. Single-verb is the existing solid pill.
  if (altMethod) {
    return (
      <span
        className={s.methodPill}
        style={{
          background: `linear-gradient(90deg, ${methodColorVar(method)} 50%, ${methodColorVar(altMethod)} 50%)`,
          fontSize: size === 'sm' ? '10px' : '11px',
          padding: size === 'sm' ? '2px 8px' : '4px 10px',
          color: tokens.colorNeutralForegroundOnBrand,
        }}
        title={`${method} (collection-valued) / ${altMethod} (single-valued) — verb depends on the chosen navigation property's cardinality`}
      >
        {method}/{altMethod}
      </span>
    );
  }
  return (
    <span
      className={s.methodPill}
      style={{
        backgroundColor: methodColorVar(method),
        fontSize: size === 'sm' ? '10px' : '11px',
        padding: size === 'sm' ? '2px 8px' : '4px 10px',
      }}
    >
      {method}
    </span>
  );
}
