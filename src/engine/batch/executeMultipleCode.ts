// ExecuteMultipleRequest (C#, SDK for .NET) from the composer's requests (#11).
//
// ExecuteMultiple is an Organization service message; the Web API has no
// ExecuteMultiple action — "Web API $batch requests are equivalent to the
// ExecuteMultiple message" (Execute batch operations using the Web API).
// So DRS sends the requests as a $batch and generates the SDK form here:
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/org-service/execute-multiple-requests
//
// Mapping (best effort, flagged where the Web API shape doesn't carry
// enough to be exact):
//   Create   → CreateRequest { Target = Entity }
//   Update   → UpdateRequest { Target = Entity with Id }
//   Delete   → DeleteRequest { Target = EntityReference }
//   Retrieve → RetrieveRequest (by id) / RetrieveMultipleRequest (QueryExpression)
//   Associate→ AssociateRequest
//   Custom POST to an unbound action → OrganizationRequest("<Name>")
// Column types come from table metadata when it's loaded (OptionSetValue,
// Money, EntityReference, …); otherwise numbers are emitted as-is with a
// TODO so the developer can confirm the type.

import type { CompiledItem } from './batchModel';

export interface CodeMetadata {
  /** `accounts` → `account`, or null when the entity set isn't known. */
  logicalOf: (entitySet: string) => string | null;
  /** Attribute type (`Picklist`, `Money`, `Lookup`, …) or null when unknown. */
  attrType: (logical: string, attribute: string) => string | null;
}

const GUID_RE = /^\{?[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\}?$/i;
const cs = (s: string) => JSON.stringify(s); // C# regular string literal is JSON-compatible for our values
const guid = (id: string) => `new Guid(${cs(id.replace(/[{}]/g, ''))})`;

function logicalFor(set: string, md: CodeMetadata): string {
  return md.logicalOf(set) ?? set.replace(/s$/, '');
}

/** `accounts(guid)` → { set, id } */
function parseRef(path: string): { set: string; id: string } | null {
  const m = /^([A-Za-z_][\w]*)\(([^)]+)\)$/.exec(path.trim());
  return m ? { set: m[1], id: m[2] } : null;
}

function valueExpr(
  logical: string,
  attr: string,
  value: unknown,
  md: CodeMetadata,
  todos: string[],
): string | null {
  if (value === null) return 'null';
  const type = md.attrType(logical, attr);
  switch (type) {
    case 'Picklist':
    case 'State':
    case 'Status':
      return `new OptionSetValue(${Number(value)})`;
    case 'MultiSelectPicklist':
      return `new OptionSetValueCollection { ${String(value)
        .split(',')
        .filter(Boolean)
        .map((v) => `new OptionSetValue(${Number(v)})`)
        .join(', ')} }`;
    case 'Money':
      return `new Money(${Number(value)}m)`;
    case 'Decimal':
      return `${Number(value)}m`;
    case 'BigInt':
      return `${Number(value)}L`;
    case 'DateTime':
      return `DateTime.Parse(${cs(String(value))})`;
    case 'Uniqueidentifier':
      return guid(String(value));
  }
  if (typeof value === 'string') return cs(value);
  if (typeof value === 'boolean') return String(value);
  if (typeof value === 'number') {
    if (!type) todos.push(`${attr}: confirm the type (int / OptionSetValue / Money / decimal)`);
    return String(value);
  }
  todos.push(`${attr}: nested value not converted`);
  return null;
}

/** Entity initializer lines for a Web API JSON body. */
function entityBody(
  logical: string,
  body: Record<string, unknown>,
  md: CodeMetadata,
  todos: string[],
): string[] {
  const lines: string[] = [];
  for (const [key, value] of Object.entries(body)) {
    if (key.endsWith('@odata.bind')) {
      let attr = key.slice(0, -'@odata.bind'.length);
      // Polymorphic binds name the target: customerid_account → customerid.
      if (!md.attrType(logical, attr) && attr.includes('_')) {
        const base = attr.slice(0, attr.lastIndexOf('_'));
        if (md.attrType(logical, base)) attr = base;
      }
      const ref =
        typeof value === 'string' ? parseRef(value.replace(/^.*\/api\/data\/v[\d.]+\//, '')) : null;
      if (typeof value === 'string' && /^\$\d+$/.test(value)) {
        todos.push(
          `${attr}: ${value} referenced an earlier request — ExecuteMultiple has no Content-ID references; set the id returned by that request`,
        );
      } else if (ref) {
        lines.push(
          `[${cs(attr)}] = new EntityReference(${cs(logicalFor(ref.set, md))}, ${guid(ref.id)}),`,
        );
      } else {
        todos.push(`${attr}: couldn't convert the @odata.bind value`);
      }
      continue;
    }
    if (key.includes('@')) continue; // annotations
    const expr = valueExpr(logical, key, value, md, todos);
    if (expr !== null) lines.push(`[${cs(key)}] = ${expr},`);
  }
  return lines;
}

function parseBody(text?: string): Record<string, unknown> {
  if (!text) return {};
  try {
    const v = JSON.parse(text) as unknown;
    return v && typeof v === 'object' && !Array.isArray(v) ? (v as Record<string, unknown>) : {};
  } catch {
    return {};
  }
}

const indent = (lines: string[], n: number) => lines.map((l) => ' '.repeat(n) + l);

function columnSet(select: string | null): string {
  if (!select) return 'new ColumnSet(true)';
  return `new ColumnSet(${select
    .split(',')
    .map((c) => cs(c.trim().replace(/^_(.+)_value$/, '$1')))
    .join(', ')})`;
}

/** One `requests.Add(...)` statement (or a comment) per composer request. */
export function requestToCSharp(c: CompiledItem, md: CodeMetadata): string[] {
  const w = c.wire;
  if (!w) return [`// $${c.contentId}: has errors — fix it in the Builder.`];
  const todos: string[] = [];
  const out: string[] = [];
  const it = c.item;
  const head = `// $${c.contentId} ${w.method} ${w.path}`;
  const finish = (expr: string[]) => {
    out.push(head, ...todos.map((t) => `// TODO ${t}`), ...expr);
    return out;
  };
  const isGuid = GUID_RE.test(it.recordId.trim());

  if (it.op === 'create' || it.op === 'update') {
    const logical = logicalFor(it.entitySet, md);
    if (it.op === 'update' && !isGuid) {
      todos.push(
        `record id ${it.recordId} isn't a GUID — use new Entity(logicalName, keyAttributes) for an alternate key`,
      );
    }
    const ctor =
      it.op === 'update' && isGuid
        ? `new Entity(${cs(logical)}, ${guid(it.recordId.trim())})`
        : `new Entity(${cs(logical)})`;
    const body = entityBody(logical, parseBody(w.body), md, todos);
    return finish([
      `requests.Add(new ${it.op === 'create' ? 'CreateRequest' : 'UpdateRequest'}`,
      '{',
      `    Target = ${ctor}`,
      '    {',
      ...indent(body, 8),
      '    },',
      '});',
    ]);
  }
  if (it.op === 'delete') {
    const logical = logicalFor(it.entitySet, md);
    if (!isGuid) todos.push(`record id ${it.recordId} isn't a GUID`);
    return finish([
      `requests.Add(new DeleteRequest { Target = new EntityReference(${cs(logical)}, ${guid(it.recordId.trim())}) });`,
    ]);
  }
  if (it.op === 'retrieve') {
    const logical = logicalFor(it.entitySet, md);
    const params = new URLSearchParams(it.query.trim().replace(/^\?/, ''));
    const select = params.get('$select');
    if (it.recordId.trim()) {
      return finish([
        'requests.Add(new RetrieveRequest',
        '{',
        `    Target = new EntityReference(${cs(logical)}, ${guid(it.recordId.trim())}),`,
        `    ColumnSet = ${columnSet(select)},`,
        '});',
      ]);
    }
    const top = params.get('$top');
    for (const k of ['$filter', '$orderby', '$expand', '$apply', 'fetchXml']) {
      if (params.get(k))
        todos.push(`translate ${k} into the QueryExpression (or use FetchExpression)`);
    }
    return finish([
      'requests.Add(new RetrieveMultipleRequest',
      '{',
      `    Query = new QueryExpression(${cs(logical)})`,
      '    {',
      `        ColumnSet = ${columnSet(select)},`,
      ...(top ? [`        TopCount = ${Number(top)},`] : []),
      '    },',
      '});',
    ]);
  }
  if (it.op === 'associate') {
    const logical = logicalFor(it.entitySet, md);
    const target = logicalFor(it.targetSet, md);
    todos.push(
      'Relationship takes the relationship schema name — usually the navigation property name',
    );
    return finish([
      'requests.Add(new AssociateRequest',
      '{',
      `    Target = new EntityReference(${cs(logical)}, ${guid(it.recordId.trim())}),`,
      `    Relationship = new Relationship(${cs(it.navProperty.trim())}),`,
      `    RelatedEntities = new EntityReferenceCollection { new EntityReference(${cs(target)}, ${guid(it.targetId.trim())}) },`,
      '});',
    ]);
  }
  // Custom: an unbound action maps to OrganizationRequest by name.
  const action = /^(?:Microsoft\.Dynamics\.CRM\.)?([A-Za-z_][\w]*)$/.exec(w.path);
  if (w.method === 'POST' && action && !md.logicalOf(action[1])) {
    const params = Object.entries(parseBody(w.body)).filter(([k]) => !k.includes('@'));
    return finish([
      `requests.Add(new OrganizationRequest(${cs(action[1])})`,
      '{',
      '    Parameters =',
      '    {',
      ...params.map(([k, v]) =>
        typeof v === 'object' && v !== null
          ? `        // TODO ${k}: convert ${JSON.stringify(v)}`
          : `        [${cs(k)}] = ${typeof v === 'string' ? cs(v) : String(v)},`,
      ),
      '    },',
      '});',
    ]);
  }
  return [
    head,
    '// No SDK request generated for this Web API call — add the matching OrganizationRequest.',
  ];
}

export function generateExecuteMultipleCSharp(
  items: CompiledItem[],
  opts: { continueOnError: boolean; returnResponses: boolean },
  md: CodeMetadata,
): string {
  const body = items.flatMap((c) => [...requestToCSharp(c, md), '']);
  return [
    '// ExecuteMultiple (SDK for .NET). Microsoft.Xrm.Sdk + Microsoft.Xrm.Sdk.Messages + Microsoft.Xrm.Sdk.Query.',
    '// Max 1000 requests per call; ExecuteMultiple is not transactional — use ExecuteTransactionRequest for all-or-nothing.',
    'var requests = new OrganizationRequestCollection();',
    '',
    ...body,
    'var executeMultiple = new ExecuteMultipleRequest',
    '{',
    '    Settings = new ExecuteMultipleSettings',
    '    {',
    `        ContinueOnError = ${opts.continueOnError},`,
    `        ReturnResponses = ${opts.returnResponses},`,
    '    },',
    '    Requests = requests,',
    '};',
    '',
    'var response = (ExecuteMultipleResponse)service.Execute(executeMultiple);',
    '',
    '// With ReturnResponses = false only faults come back. RequestIndex is 0-based.',
    'foreach (var item in response.Responses)',
    '{',
    '    if (item.Fault != null)',
    '        Console.WriteLine($"Request {item.RequestIndex} failed: {item.Fault.Message}");',
    '    else if (item.Response != null)',
    '        Console.WriteLine($"Request {item.RequestIndex}: {item.Response.ResponseName}");',
    '}',
    `Console.WriteLine($"IsFaulted: {response.IsFaulted}");`,
  ].join('\n');
}
