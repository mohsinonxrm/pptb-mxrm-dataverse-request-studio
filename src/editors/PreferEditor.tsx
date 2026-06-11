// Prefer header builder. Spec: §16.
import {
  Switch,
  Field,
  SpinButton,
  Caption1,
  tokens,
  Checkbox,
  Body1,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
} from '@fluentui/react-components';
import { Settings20Filled } from '@fluentui/react-icons';
import { PaneHead } from './PaneHead';
import { useHostSession } from '../host/HostContext';
import type { RequestGroup } from '../registry/requestTypes';

export interface PreferSpec {
  /** odata.maxpagesize=N (mutually exclusive with $top) */
  maxpagesize: number | null;
  /** Annotation modes */
  formattedValues: boolean;
  lookupLogicalNames: boolean;
  associatedNavProperty: boolean;
  totalRecordCount: boolean;
  /** Wildcard catch-all (overrides individual ones) */
  allAnnotations: boolean;
  /** return=representation — only relevant on writes; included for completeness */
  returnRepresentation: boolean;
  /** include-strict-validation, etc. could go here later */
}

export const emptyPrefer = (): PreferSpec => ({
  maxpagesize: null,
  formattedValues: true,
  lookupLogicalNames: false,
  associatedNavProperty: false,
  totalRecordCount: false,
  allAnnotations: false,
  returnRepresentation: false,
});

export function PreferEditor({
  spec,
  setSpec,
  group = 'read',
}: {
  spec: PreferSpec;
  setSpec: (s: PreferSpec) => void;
  group?: RequestGroup;
}) {
  const set = <K extends keyof PreferSpec>(k: K, v: PreferSpec[K]) => setSpec({ ...spec, [k]: v });
  const host = useHostSession();

  return (
    <div>
      <PaneHead
        icon={Settings20Filled}
        title="Prefer"
        sub="HTTP Prefer-header options that change the response shape (annotations, pagination)."
        group={group}
      />

      {/* PPTB's `dataverseAPI.queryData` builds the fetch internally and does
          NOT forward custom request headers — Prefer, If-Match, MSCRM.*
          values are dropped. The host substitutes its own defaults (which
          DO include common annotations like FormattedValue / lookup logical
          name), so basic display works. But user-tuned values — small page
          sizes for testing, return=representation, etc. — don't apply on
          Execute. They still appear in the Code tab samples (curl / Postman
          / fetch) for use against the Web API directly. */}
      {host.embedded && (
        <MessageBar layout="multiline" intent="info" style={{ marginBottom: 14, maxWidth: 720 }}>
          <MessageBarBody>
            <MessageBarTitle>Prefer headers aren't honored by Execute inside PPTB.</MessageBarTitle>
            The host builds the request itself and ignores custom headers you set here. Inside PPTB,
            Dataverse uses its server defaults: page size <strong>5,000</strong>, no{' '}
            <code>return=representation</code> echo on writes. Common annotations (FormattedValue,
            lookuplogicalname) are still emitted by the host's own default Prefer, so the grid
            renders correctly.{' '}
            <strong>The values you set here are emitted into the Code tab</strong> — export to
            Postman / curl / fetch to test the full request shape (including small page sizes for
            pagination testing).
          </MessageBarBody>
        </MessageBar>
      )}

      <div style={{ display: 'flex', flexDirection: 'column', gap: 18, maxWidth: 720 }}>
        <section>
          <Caption1
            style={{
              display: 'block',
              marginBottom: 6,
              fontWeight: 600,
              color: tokens.colorNeutralForeground2,
            }}
          >
            Pagination
          </Caption1>
          <Field
            label="odata.maxpagesize"
            hint="Server-emitted page size — issues @odata.nextLink for the next page. Mutually exclusive with $top."
          >
            <SpinButton
              value={spec.maxpagesize ?? 0}
              min={0}
              max={5000}
              step={10}
              onChange={(_, d) => {
                const v = d.value ?? Number(d.displayValue ?? 0);
                set('maxpagesize', v === 0 ? null : v);
              }}
            />
          </Field>
        </section>

        <section>
          <Caption1
            style={{
              display: 'block',
              marginBottom: 6,
              fontWeight: 600,
              color: tokens.colorNeutralForeground2,
            }}
          >
            Annotations{' '}
            <span
              style={{
                fontWeight: 400,
                color: tokens.colorNeutralForeground3,
                fontFamily: tokens.fontFamilyMonospace,
                fontSize: 10,
              }}
            >
              · odata.include-annotations
            </span>
          </Caption1>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
            <Switch
              checked={spec.allAnnotations}
              onChange={(_, d) => set('allAnnotations', d.checked)}
              label={
                <Body1>
                  All annotations (<code>"*"</code>)
                </Body1>
              }
            />
            <div
              style={{
                marginLeft: 24,
                display: 'flex',
                flexDirection: 'column',
                gap: 2,
                opacity: spec.allAnnotations ? 0.45 : 1,
                pointerEvents: spec.allAnnotations ? 'none' : 'auto',
              }}
            >
              <Checkbox
                checked={spec.formattedValues}
                onChange={(_, d) => set('formattedValues', !!d.checked)}
                label={
                  <span>
                    <strong>OData.Community.Display.V1.FormattedValue</strong> — localized display
                    strings for choices, currency, dates, lookups
                  </span>
                }
              />
              <Checkbox
                checked={spec.lookupLogicalNames}
                onChange={(_, d) => set('lookupLogicalNames', !!d.checked)}
                label={
                  <span>
                    <strong>Microsoft.Dynamics.CRM.lookuplogicalname</strong> — disambiguates
                    polymorphic lookups (ownerid → systemuser/team)
                  </span>
                }
              />
              <Checkbox
                checked={spec.associatedNavProperty}
                onChange={(_, d) => set('associatedNavProperty', !!d.checked)}
                label={
                  <span>
                    <strong>Microsoft.Dynamics.CRM.associatednavigationproperty</strong> — gives the
                    nav-prop name for follow-up $expand
                  </span>
                }
              />
              <Checkbox
                checked={spec.totalRecordCount}
                onChange={(_, d) => set('totalRecordCount', !!d.checked)}
                label={
                  <span>
                    <strong>Microsoft.Dynamics.CRM.totalrecordcount[…limitexceeded]</strong> —
                    disambiguate "is this 5,000 or really 5,000?"
                  </span>
                }
              />
            </div>
          </div>
        </section>

        <section>
          <Caption1
            style={{
              display: 'block',
              marginBottom: 6,
              fontWeight: 600,
              color: tokens.colorNeutralForeground2,
            }}
          >
            Other
          </Caption1>
          <Switch
            checked={spec.returnRepresentation}
            onChange={(_, d) => set('returnRepresentation', d.checked)}
            label={
              <span>
                return=representation{' '}
                <Caption1 style={{ marginLeft: 4, color: tokens.colorNeutralForeground3 }}>
                  (write-side; echoes the row back)
                </Caption1>
              </span>
            }
          />
        </section>
      </div>
    </div>
  );
}

export function preferToHeaderValues(spec: PreferSpec): string[] {
  const out: string[] = [];
  if (spec.maxpagesize != null && spec.maxpagesize > 0) {
    out.push(`odata.maxpagesize=${spec.maxpagesize}`);
  }
  if (spec.allAnnotations) {
    out.push('odata.include-annotations="*"');
  } else {
    const annos: string[] = [];
    if (spec.formattedValues) annos.push('OData.Community.Display.V1.FormattedValue');
    if (spec.lookupLogicalNames) annos.push('Microsoft.Dynamics.CRM.lookuplogicalname');
    if (spec.associatedNavProperty)
      annos.push('Microsoft.Dynamics.CRM.associatednavigationproperty');
    if (spec.totalRecordCount)
      annos.push(
        'Microsoft.Dynamics.CRM.totalrecordcount,Microsoft.Dynamics.CRM.totalrecordcountlimitexceeded',
      );
    if (annos.length) out.push(`odata.include-annotations="${annos.join(',')}"`);
  }
  if (spec.returnRepresentation) out.push('return=representation');
  return out;
}

export function preferToHeaderString(spec: PreferSpec): string | null {
  const vals = preferToHeaderValues(spec);
  return vals.length ? vals.join(', ') : null;
}
