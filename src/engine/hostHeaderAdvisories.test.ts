import { beforeEach, describe, expect, it, vi } from 'vitest';
import {
  detectUnsentHeaderAdvisories,
  isConcurrencyEnforcementRequired,
  type HeaderSensitiveOperation,
} from './hostHeaderAdvisories';
import { defaultBypassOptions, type ConcurrencyMode } from '../state/writeState';

// The bridge decides host capabilities once at module load; mock it so each
// test can pick the host it runs against.
const host = vi.hoisted(() => ({ embedded: true, additionalHeaders: false }));
vi.mock('../host/pptbBridge', () => ({
  isEmbedded: () => host.embedded,
  hostSupportsAdditionalHeaders: () => host.additionalHeaders,
}));

const ops: HeaderSensitiveOperation[] = ['create', 'update', 'upsert', 'delete', 'merge'];

beforeEach(() => {
  host.embedded = true;
  host.additionalHeaders = false;
});

describe('isConcurrencyEnforcementRequired', () => {
  it.each(ops)('%s: none never needs enforcement; etag and create-only always do', (op) => {
    expect(isConcurrencyEnforcementRequired(op, { kind: 'none' })).toBe(false);
    expect(isConcurrencyEnforcementRequired(op, { kind: 'etag', etag: 'W/"1"' })).toBe(true);
    expect(isConcurrencyEnforcementRequired(op, { kind: 'create-only' })).toBe(true);
  });

  it('update-only only matters for Upsert (it is the Update default)', () => {
    const updateOnly: ConcurrencyMode = { kind: 'update-only' };
    expect(isConcurrencyEnforcementRequired('upsert', updateOnly)).toBe(true);
    expect(isConcurrencyEnforcementRequired('update', updateOnly)).toBe(false);
    expect(isConcurrencyEnforcementRequired('delete', updateOnly)).toBe(false);
  });
});

describe('detectUnsentHeaderAdvisories', () => {
  const bypassOn = { ...defaultBypassOptions(), businessLogic: 'sync' as const };

  it('blocks a precondition the host would drop, focused on the precondition node', () => {
    const out = detectUnsentHeaderAdvisories({
      op: 'upsert',
      concurrency: { kind: 'create-only' },
    });
    expect(out).toHaveLength(1);
    expect(out[0]).toMatchObject({
      id: 'host-precondition-unsent',
      severity: 'error',
      source: 'header',
      focusNode: 'precondition',
    });
  });

  it('blocks active bypass options and the legacy Delete plug-in flag', () => {
    expect(
      detectUnsentHeaderAdvisories({ op: 'create', bypass: bypassOn }).map((a) => a.id),
    ).toEqual(['host-bypass-unsent']);
    expect(
      detectUnsentHeaderAdvisories({
        op: 'create',
        bypass: { ...defaultBypassOptions(), suppressFlows: true },
      }).map((a) => a.id),
    ).toEqual(['host-bypass-unsent']);
    expect(
      detectUnsentHeaderAdvisories({ op: 'delete', bypassCustomPlugins: true }).map((a) => a.id),
    ).toEqual(['host-bypass-unsent']);
  });

  it('reports both and honours custom focus nodes', () => {
    const out = detectUnsentHeaderAdvisories({
      op: 'update',
      concurrency: { kind: 'etag', etag: 'W/"7"' },
      bypass: bypassOn,
      preconditionFocus: 'conc',
      bypassFocus: 'byp',
    });
    expect(out.map((a) => [a.id, a.focusNode])).toEqual([
      ['host-precondition-unsent', 'conc'],
      ['host-bypass-unsent', 'byp'],
    ]);
  });

  it('stays quiet for the defaults (no precondition, bypass off, Update If-Match: *)', () => {
    expect(
      detectUnsentHeaderAdvisories({
        op: 'update',
        concurrency: { kind: 'update-only' },
        bypass: defaultBypassOptions(),
      }),
    ).toEqual([]);
  });

  it('stays quiet on PPTB 1.2.7+ hosts that send additionalHeaders (#43)', () => {
    host.additionalHeaders = true;
    expect(
      detectUnsentHeaderAdvisories({
        op: 'upsert',
        concurrency: { kind: 'create-only' },
        bypass: bypassOn,
        bypassCustomPlugins: true,
      }),
    ).toEqual([]);
  });

  it('stays quiet outside PPTB (nothing executes there)', () => {
    host.embedded = false;
    expect(
      detectUnsentHeaderAdvisories({ op: 'upsert', concurrency: { kind: 'create-only' } }),
    ).toEqual([]);
  });
});
