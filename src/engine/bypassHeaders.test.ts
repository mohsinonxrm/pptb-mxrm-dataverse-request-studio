// Bypass pane → MSCRM.* headers: the new vs legacy business-logic header,
// step ids, the Power Automate flow bypass, and how the pane's rows merge
// with headers the user typed (#235).
import { describe, expect, it } from 'vitest';
import {
  STEP_IDS_DEFAULT_LIMIT,
  STEP_IDS_MAX_LIMIT,
  applyBypassToHeaders,
  composeBypassHeaders,
} from './bypassHeaders';
import { toWireHeaders } from './requestHeaderPolicy';
import { defaultBypassOptions, type BypassOptions } from '../state/writeState';
import type { HeaderItem } from '../editors/HeadersEditor';

const bypass = (b: Partial<BypassOptions>): BypassOptions => ({ ...defaultBypassOptions(), ...b });
const wire = (items: HeaderItem[]) => items.map((h) => [h.name, h.value]);
const header = (id: string, name: string, value: string): HeaderItem => ({
  id,
  name,
  value,
  enabled: true,
});

const STEP_A = '00000000-0000-0000-0000-00000000000a';
const STEP_B = '00000000-0000-0000-0000-00000000000b';

describe('composeBypassHeaders', () => {
  it('sends nothing by default', () => {
    expect(composeBypassHeaders(defaultBypassOptions())).toEqual([]);
  });

  it('sync: the BypassBusinessLogicExecution header with CustomSync', () => {
    expect(wire(composeBypassHeaders(bypass({ businessLogic: 'sync' })))).toEqual([
      ['MSCRM.BypassBusinessLogicExecution', 'CustomSync'],
    ]);
  });

  it('sync with the legacy option: BypassCustomPluginExecution instead', () => {
    expect(
      wire(composeBypassHeaders(bypass({ businessLogic: 'sync', useLegacyHeader: true }))),
    ).toEqual([['MSCRM.BypassCustomPluginExecution', 'true']]);
  });

  it('async and both use the new header; the legacy option only applies to sync', () => {
    expect(
      wire(composeBypassHeaders(bypass({ businessLogic: 'async', useLegacyHeader: true }))),
    ).toEqual([['MSCRM.BypassBusinessLogicExecution', 'CustomAsync']]);
    expect(
      wire(composeBypassHeaders(bypass({ businessLogic: 'both', useLegacyHeader: true }))),
    ).toEqual([['MSCRM.BypassBusinessLogicExecution', 'CustomSync,CustomAsync']]);
  });

  it('steps: the step ids comma-separated, skipping blank entries', () => {
    expect(
      wire(composeBypassHeaders(bypass({ businessLogic: 'steps', stepIds: [STEP_A, '', STEP_B] }))),
    ).toEqual([['MSCRM.BypassBusinessLogicExecutionStepIds', `${STEP_A},${STEP_B}`]]);
  });

  it('steps with no ids sends no header', () => {
    expect(composeBypassHeaders(bypass({ businessLogic: 'steps', stepIds: [] }))).toEqual([]);
    expect(composeBypassHeaders(bypass({ businessLogic: 'steps', stepIds: ['', ''] }))).toEqual([]);
  });

  it('step ids are ignored unless the steps option is chosen', () => {
    expect(
      wire(composeBypassHeaders(bypass({ businessLogic: 'sync', stepIds: [STEP_A] }))),
    ).toEqual([['MSCRM.BypassBusinessLogicExecution', 'CustomSync']]);
  });

  it('the steps hint states the default and maximum step-id caps', () => {
    const [h] = composeBypassHeaders(bypass({ businessLogic: 'steps', stepIds: [STEP_A] }));
    expect(h.hint).toContain(
      `Default cap is ${STEP_IDS_DEFAULT_LIMIT} (max ${STEP_IDS_MAX_LIMIT})`,
    );
  });

  it('suppressFlows: SuppressCallbackRegistrationExpanderJob, alone or after the business-logic header', () => {
    expect(wire(composeBypassHeaders(bypass({ suppressFlows: true })))).toEqual([
      ['MSCRM.SuppressCallbackRegistrationExpanderJob', 'true'],
    ]);
    expect(
      wire(
        composeBypassHeaders(
          bypass({ businessLogic: 'sync', useLegacyHeader: true, suppressFlows: true }),
        ),
      ),
    ).toEqual([
      ['MSCRM.BypassCustomPluginExecution', 'true'],
      ['MSCRM.SuppressCallbackRegistrationExpanderJob', 'true'],
    ]);
  });

  it('every row is an enabled built-in with a hint', () => {
    const rows = composeBypassHeaders(bypass({ businessLogic: 'both', suppressFlows: true }));
    expect(rows).toHaveLength(2);
    for (const r of rows) {
      expect(r).toMatchObject({ enabled: true, builtin: true });
      expect(r.hint).toBeTruthy();
    }
  });
});

describe('applyBypassToHeaders', () => {
  const user = [header('h1', 'Prefer', 'return=representation'), header('h2', 'X-Custom', '1')];

  it("keeps the user's headers in order and appends the bypass rows", () => {
    const out = applyBypassToHeaders(user, bypass({ businessLogic: 'async', suppressFlows: true }));
    expect(wire(out)).toEqual([
      ['Prefer', 'return=representation'],
      ['X-Custom', '1'],
      ['MSCRM.BypassBusinessLogicExecution', 'CustomAsync'],
      ['MSCRM.SuppressCallbackRegistrationExpanderJob', 'true'],
    ]);
  });

  it('replaces earlier bypass rows rather than adding to them', () => {
    const once = applyBypassToHeaders(user, bypass({ businessLogic: 'sync' }));
    const twice = applyBypassToHeaders(once, bypass({ businessLogic: 'sync' }));
    expect(twice).toEqual(once);
    const switched = applyBypassToHeaders(
      once,
      bypass({ businessLogic: 'sync', useLegacyHeader: true }),
    );
    expect(wire(switched)).toEqual([
      ['Prefer', 'return=representation'],
      ['X-Custom', '1'],
      ['MSCRM.BypassCustomPluginExecution', 'true'],
    ]);
  });

  it('turning the bypass off removes its rows and leaves the user headers', () => {
    const on = applyBypassToHeaders(user, bypass({ businessLogic: 'both', suppressFlows: true }));
    expect(applyBypassToHeaders(on, defaultBypassOptions())).toEqual(user);
  });

  it('does not change the list it was given', () => {
    const input = [...user];
    applyBypassToHeaders(input, bypass({ businessLogic: 'sync' }));
    expect(input).toEqual(user);
  });

  it('keeps a header the user typed with the same name, ahead of the bypass row', () => {
    const typed = header('h3', 'MSCRM.BypassBusinessLogicExecution', 'CustomAsync');
    const out = applyBypassToHeaders([typed], bypass({ businessLogic: 'sync' }));
    expect(out[0]).toBe(typed);
    expect(wire(out)).toEqual([
      ['MSCRM.BypassBusinessLogicExecution', 'CustomAsync'],
      ['MSCRM.BypassBusinessLogicExecution', 'CustomSync'],
    ]);
  });

  it('so on the wire the bypass pane wins over a same-named typed header, in any case', () => {
    const typed = header('h3', 'mscrm.bypassbusinesslogicexecution', 'CustomAsync');
    const out = applyBypassToHeaders([typed], bypass({ businessLogic: 'sync' }));
    const record = Object.fromEntries(out.filter((h) => h.enabled).map((h) => [h.name, h.value]));
    expect(toWireHeaders(record)).toEqual({ 'MSCRM.BypassBusinessLogicExecution': 'CustomSync' });
  });

  it('with the bypass off, a typed bypass header is left alone', () => {
    const typed = header('h3', 'MSCRM.SuppressCallbackRegistrationExpanderJob', 'true');
    expect(applyBypassToHeaders([typed], defaultBypassOptions())).toEqual([typed]);
  });
});
