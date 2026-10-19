// Workflow XAML → declared input arguments (#4).
//
// A process's `workflow.xaml` declares its arguments as `x:Property`
// members of the root Activity:
//
//   <x:Members>
//     <x:Property Name="InputEntities" Type="InArgument(scg:IDictionary(x:String, mxs:Entity))" />
//     <x:Property Name="Reason" Type="InArgument(x:String)">
//       <x:Property.Attributes>
//         <mxsw:ArgumentRequiredAttribute Value="True" />
//         <mxsw:ArgumentDescriptionAttribute Value="Why" />
//       </x:Property.Attributes>
//     </x:Property>
//   </x:Members>
//
// `InputEntities` / `CreatedEntities` are plumbing every classic workflow
// carries (the runtime fills them), so they're not user inputs. Parsing is
// best effort: malformed XAML or an unexpected shape yields no arguments
// rather than an error.

const XAML_NS = 'http://schemas.microsoft.com/winfx/2006/xaml';

/** Runtime-supplied arguments present on every classic workflow. */
const SYSTEM_ARGUMENTS = new Set(['InputEntities', 'CreatedEntities']);

export interface WorkflowArgument {
  name: string;
  direction: 'In' | 'InOut';
  /** The XAML type inside InArgument(...), e.g. `x:String`, `mxs:EntityReference`. */
  type: string;
  required: boolean;
  description: string | null;
}

export function parseWorkflowInArguments(xaml: string | null | undefined): WorkflowArgument[] {
  if (!xaml || typeof DOMParser === 'undefined') return [];
  let doc: Document;
  try {
    doc = new DOMParser().parseFromString(xaml, 'application/xml');
  } catch {
    return [];
  }
  if (doc.getElementsByTagName('parsererror').length > 0) return [];

  const out: WorkflowArgument[] = [];
  const props = doc.getElementsByTagNameNS(XAML_NS, 'Property');
  for (const el of Array.from(props)) {
    const name = el.getAttribute('Name');
    const m = /^(In|InOut)Argument\((.*)\)$/.exec((el.getAttribute('Type') ?? '').trim());
    if (!name || !m || SYSTEM_ARGUMENTS.has(name)) continue;
    let required = false;
    let description: string | null = null;
    for (const attr of Array.from(el.getElementsByTagName('*'))) {
      const value = attr.getAttribute('Value');
      if (attr.localName === 'ArgumentRequiredAttribute')
        required = value?.toLowerCase() === 'true';
      if (attr.localName === 'ArgumentDescriptionAttribute' && value) description = value;
    }
    out.push({ name, direction: m[1] as 'In' | 'InOut', type: m[2].trim(), required, description });
  }
  return out;
}

/** Short display label for a XAML argument type: `x:String` → `String`. */
export function workflowArgumentTypeLabel(type: string): string {
  return type.replace(/\b[a-z]+:/g, '');
}
