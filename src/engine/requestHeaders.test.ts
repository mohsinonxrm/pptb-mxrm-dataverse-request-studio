// One header composition for Execute, the Code tab and the exports (#244).
import { describe, expect, it } from 'vitest';
import { composeHeaderItems, composeHeaders, type HeaderBits } from './requestHeaders';
import { composeHeaders as exportHeaders } from './savedToCollection';
import { composePartHeaders } from './batch/savedToBatch';
import { COUNT_ONLY_PREFER } from './urlBuilder';
import { defaultWriteHeaders, headerItemsToObject } from '../editors/HeadersEditor';
import { emptyPrefer } from '../editors/PreferEditor';
import { defaultBypassOptions } from '../state/writeState';

const names = (bits: HeaderBits) =>
  composeHeaderItems(bits)
    .filter((h) => h.enabled)
    .map((h) => `${h.name}: ${h.value}`);

describe('composeHeaderItems', () => {
  it('adds the precondition after the rows, then the bypass headers', () => {
    expect(
      names({
        headers: [{ id: 'a', name: 'X-Mine', value: '1', enabled: true }],
        concurrency: { kind: 'etag', etag: ' W/"5" ' },
        bypass: { ...defaultBypassOptions(), businessLogic: 'sync', suppressFlows: true },
      }),
    ).toEqual([
      'X-Mine: 1',
      'If-Match: W/"5"',
      'MSCRM.BypassBusinessLogicExecution: CustomSync',
      'MSCRM.SuppressCallbackRegistrationExpanderJob: true',
    ]);
  });

  it("Create's duplicate detection switch drives its row", () => {
    const row = (on: boolean) =>
      composeHeaderItems({ headers: defaultWriteHeaders(), duplicateDetection: on }).find(
        (h) => h.name === 'MSCRM.SuppressDuplicateDetection',
      );
    expect(row(true)).toMatchObject({ enabled: true, value: 'false' });
    expect(row(false)).toMatchObject({ enabled: false });
  });

  it("Merge's skip adds SuppressDuplicateDetection: true", () => {
    expect(names({ suppressDuplicateDetection: true })).toEqual([
      'MSCRM.SuppressDuplicateDetection: true',
    ]);
  });

  it("Delete's legacy switch sends the legacy plug-in bypass only when the Bypass pane is off", () => {
    expect(names({ bypass: defaultBypassOptions(), bypassCustomPlugins: true })).toEqual([
      'MSCRM.BypassCustomPluginExecution: true',
    ]);
    expect(
      names({
        bypass: { ...defaultBypassOptions(), businessLogic: 'async' },
        bypassCustomPlugins: true,
      }),
    ).toEqual(['MSCRM.BypassBusinessLogicExecution: CustomAsync']);
  });

  it('a composed header wins over a row of the same name, as Execute sends it', () => {
    const mine = (name: string, value: string) => ({ id: name, name, value, enabled: true });
    expect(
      composeHeaders({
        headers: [mine('If-Match', 'W/"old"')],
        concurrency: { kind: 'etag', etag: 'W/"new"' },
      })['If-Match'],
    ).toBe('W/"new"');
    expect(
      composeHeaders({
        headers: [mine('MSCRM.SuppressDuplicateDetection', 'false')],
        suppressDuplicateDetection: true,
      })['MSCRM.SuppressDuplicateDetection'],
    ).toBe('true');
    expect(
      composeHeaders({
        headers: [mine('MSCRM.BypassBusinessLogicExecution', 'CustomAsync')],
        bypass: { ...defaultBypassOptions(), businessLogic: 'sync' },
      })['MSCRM.BypassBusinessLogicExecution'],
    ).toBe('CustomSync');
  });

  it('If-Match: * and If-None-Match: * for update-only and create-only', () => {
    expect(names({ concurrency: { kind: 'update-only' } })).toEqual(['If-Match: *']);
    expect(names({ concurrency: { kind: 'create-only' } })).toEqual(['If-None-Match: *']);
    expect(names({ concurrency: { kind: 'none' } })).toEqual([]);
  });
});

describe('Execute, the Code tab and the exports send the same headers', () => {
  const states: [string, HeaderBits][] = [
    [
      'Update',
      {
        headers: defaultWriteHeaders(),
        prefer: { ...emptyPrefer(), returnRepresentation: true },
        concurrency: { kind: 'update-only' },
        bypass: { ...defaultBypassOptions(), businessLogic: 'both' },
      },
    ],
    [
      'Delete',
      {
        headers: defaultWriteHeaders(),
        concurrency: { kind: 'etag', etag: 'W/"9"' },
        bypass: defaultBypassOptions(),
        bypassCustomPlugins: true,
      },
    ],
    [
      'Create',
      { headers: defaultWriteHeaders(), duplicateDetection: true, bypass: defaultBypassOptions() },
    ],
    ['Merge', { headers: defaultWriteHeaders(), suppressDuplicateDetection: true }],
  ];

  it.each(states)('%s', (_, s) => {
    // What a mode builds for the Code tab (and, without Prefer, for Execute).
    const mode = headerItemsToObject(
      composeHeaderItems(s),
      s.prefer ? (composeHeaders({ prefer: s.prefer }).Prefer ?? null) : null,
    );
    expect(exportHeaders(s)).toEqual(mode);
    // A $batch part leaves out only what the host sets itself.
    const part = composePartHeaders(s);
    for (const [k, v] of Object.entries(part)) expect(mode[k]).toBe(v);
    const left = Object.keys(mode)
      .filter((k) => !(k in part))
      .map((k) => k.toLowerCase());
    for (const k of left)
      expect(['accept', 'content-type', 'odata-version', 'odata-maxversion']).toContain(k);
  });

  it('Count only sends its own Prefer', () => {
    expect(
      composeHeaders({ countOnly: true, prefer: { ...emptyPrefer(), maxpagesize: 50 } }).Prefer,
    ).toBe(COUNT_ONLY_PREFER);
  });
});
