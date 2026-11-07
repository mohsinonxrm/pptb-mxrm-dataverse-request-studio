// Execute Workflow's input arguments, read from the process XAML (#4, #235):
// In / InOut arguments and their types, required and description
// attributes, the runtime-supplied arguments, namespaces, and XAML that
// can't be read.
import { describe, expect, it } from 'vitest';
import { parseWorkflowInArguments, workflowArgumentTypeLabel } from './workflowXaml';

const X = 'http://schemas.microsoft.com/winfx/2006/xaml';
const NS = [
  'xmlns="http://schemas.microsoft.com/netfx/2009/xaml/activities"',
  `xmlns:x="${X}"`,
  'xmlns:mxs="clr-namespace:Microsoft.Xrm.Sdk;assembly=Microsoft.Xrm.Sdk"',
  'xmlns:mxsw="clr-namespace:Microsoft.Xrm.Sdk.Workflow;assembly=Microsoft.Xrm.Sdk.Workflow"',
  'xmlns:scg="clr-namespace:System.Collections.Generic;assembly=mscorlib"',
].join(' ');

const activity = (members: string, ns = NS) =>
  `<Activity x:Class="XrmWorkflow" ${ns}><x:Members>${members}</x:Members><Sequence /></Activity>`;

const attrs = (...inner: string[]) =>
  `<x:Property.Attributes>${inner.join('')}</x:Property.Attributes>`;
const required = (v: string) => `<mxsw:ArgumentRequiredAttribute Value="${v}" />`;
const description = (v: string) => `<mxsw:ArgumentDescriptionAttribute Value="${v}" />`;

const CLASSIC_WORKFLOW = activity(
  [
    '<x:Property Name="InputEntities" Type="InArgument(scg:IDictionary(x:String, mxs:Entity))" />',
    '<x:Property Name="CreatedEntities" Type="InArgument(scg:IDictionary(x:String, mxs:Entity))" />',
    `<x:Property Name="Reason" Type="InArgument(x:String)">${attrs(
      required('True'),
      description('Why the case is closed'),
      '<mxsw:ArgumentDirectionAttribute Value="Input" />',
    )}</x:Property>`,
    `<x:Property Name="Target" Type="InArgument(mxs:EntityReference)">${attrs(
      required('False'),
      '<mxsw:ArgumentEntityAttribute Value="account" />',
    )}</x:Property>`,
    '<x:Property Name="Counter" Type="InOutArgument(x:Int32)" />',
    '<x:Property Name="Result" Type="OutArgument(x:String)" />',
  ].join(''),
);

describe('parseWorkflowInArguments', () => {
  it("lists a classic workflow's input arguments, without the runtime ones", () => {
    expect(parseWorkflowInArguments(CLASSIC_WORKFLOW)).toEqual([
      {
        name: 'Reason',
        direction: 'In',
        type: 'x:String',
        required: true,
        description: 'Why the case is closed',
      },
      {
        name: 'Target',
        direction: 'In',
        type: 'mxs:EntityReference',
        required: false,
        description: null,
      },
      { name: 'Counter', direction: 'InOut', type: 'x:Int32', required: false, description: null },
    ]);
  });

  it('leaves out Out arguments and properties that are not arguments', () => {
    const xaml = activity(
      [
        '<x:Property Name="Result" Type="OutArgument(x:String)" />',
        '<x:Property Name="Plain" Type="x:String" />',
        '<x:Property Name="NoType" />',
        '<x:Property Type="InArgument(x:String)" />',
      ].join(''),
    );
    expect(parseWorkflowInArguments(xaml)).toEqual([]);
  });

  it('keeps a nested generic type whole, and trims spaces around the type', () => {
    const xaml = activity(
      [
        '<x:Property Name="Map" Type="InArgument(scg:IDictionary(x:String, mxs:Entity))" />',
        '<x:Property Name="Padded" Type="  InArgument( mxs:Money )  " />',
      ].join(''),
    );
    expect(parseWorkflowInArguments(xaml).map((a) => [a.name, a.type])).toEqual([
      ['Map', 'scg:IDictionary(x:String, mxs:Entity)'],
      ['Padded', 'mxs:Money'],
    ]);
  });

  it('reads the required flag case-insensitively', () => {
    const xaml = activity(
      [
        `<x:Property Name="A" Type="InArgument(x:String)">${attrs(required('true'))}</x:Property>`,
        `<x:Property Name="B" Type="InArgument(x:String)">${attrs(required('TRUE'))}</x:Property>`,
        `<x:Property Name="C" Type="InArgument(x:String)">${attrs(required('False'))}</x:Property>`,
        `<x:Property Name="D" Type="InArgument(x:String)">${attrs(required('yes'))}</x:Property>`,
      ].join(''),
    );
    expect(parseWorkflowInArguments(xaml).map((a) => [a.name, a.required])).toEqual([
      ['A', true],
      ['B', true],
      ['C', false],
      ['D', false],
    ]);
  });

  it('treats an empty description as none', () => {
    const xaml = activity(
      `<x:Property Name="A" Type="InArgument(x:String)">${attrs(description(''))}</x:Property>`,
    );
    expect(parseWorkflowInArguments(xaml)[0].description).toBeNull();
  });

  it("does not take one argument's attributes for another", () => {
    const xaml = activity(
      [
        `<x:Property Name="A" Type="InArgument(x:String)">${attrs(required('True'), description('first'))}</x:Property>`,
        '<x:Property Name="B" Type="InArgument(x:String)" />',
      ].join(''),
    );
    expect(parseWorkflowInArguments(xaml)[1]).toMatchObject({
      name: 'B',
      required: false,
      description: null,
    });
  });

  it('matches the XAML namespace, whatever prefix it is bound to', () => {
    const ns = NS.replace('xmlns:x=', 'xmlns:xaml=');
    const xaml = `<Activity ${ns}><xaml:Members><xaml:Property Name="Reason" Type="InArgument(x:String)" /></xaml:Members></Activity>`;
    expect(parseWorkflowInArguments(xaml).map((a) => a.name)).toEqual(['Reason']);
  });

  it('ignores Property elements in other namespaces', () => {
    const xaml = activity(
      '<x:Property Name="Real" Type="InArgument(x:String)" />',
      `${NS} xmlns:other="urn:other"`,
    ).replace(
      '<Sequence />',
      '<other:Property Name="Fake" Type="InArgument(x:String)" /><Property Name="Unprefixed" Type="InArgument(x:String)" />',
    );
    expect(parseWorkflowInArguments(xaml).map((a) => a.name)).toEqual(['Real']);
  });

  it('returns no arguments for a workflow with none', () => {
    expect(parseWorkflowInArguments(activity(''))).toEqual([]);
    expect(parseWorkflowInArguments(`<Activity ${NS}><Sequence /></Activity>`)).toEqual([]);
  });

  it('returns no arguments, rather than throwing, for missing or malformed XAML', () => {
    expect(parseWorkflowInArguments(null)).toEqual([]);
    expect(parseWorkflowInArguments(undefined)).toEqual([]);
    expect(parseWorkflowInArguments('')).toEqual([]);
    expect(parseWorkflowInArguments('not xml at all')).toEqual([]);
    expect(
      parseWorkflowInArguments(CLASSIC_WORKFLOW.replace('</x:Members>', '</x:Member>')),
    ).toEqual([]);
    // An undeclared prefix is a namespace error.
    expect(
      parseWorkflowInArguments(
        '<Activity><x:Property Name="A" Type="InArgument(x:String)" /></Activity>',
      ),
    ).toEqual([]);
  });
});

describe('workflowArgumentTypeLabel', () => {
  it('drops namespace prefixes', () => {
    expect(workflowArgumentTypeLabel('x:String')).toBe('String');
    expect(workflowArgumentTypeLabel('mxs:EntityReference')).toBe('EntityReference');
  });

  it('drops them inside generic types too', () => {
    expect(workflowArgumentTypeLabel('scg:IDictionary(x:String, mxs:Entity)')).toBe(
      'IDictionary(String, Entity)',
    );
  });

  it('leaves a type without a prefix as it is', () => {
    expect(workflowArgumentTypeLabel('Boolean')).toBe('Boolean');
  });
});
