// Code tab output for the $batch composer (#9).
//
//   http   — the raw POST /$batch request: outer headers + multipart body
//   fetch  — JavaScript that posts that body (run with your own token)
//   pptb   — the dataverseAPI.executeBatch / executeTransaction calls DRS
//            makes for this batch, in order (one per change set / run of
//            independent requests)

import { API_BASE, type CompiledItem, type DispatchSegment } from './batchModel';
import { CRLF, encodeBatch, type EncodableGroup } from './multipart';
import { toHostRequest } from './batchExecutor';

export type BatchCodeFormat = 'http' | 'fetch' | 'pptb';

export const BATCH_CODE_LABELS: Record<BatchCodeFormat, string> = {
  http: 'HTTP ($batch)',
  fetch: 'fetch (JavaScript)',
  pptb: 'PPTB dataverseAPI',
};

export interface BatchCodeInput {
  /** Compiled items with a wire request (invalid items are skipped). */
  items: CompiledItem[];
  plan: DispatchSegment[];
  continueOnError: boolean;
  /** `https://contoso.crm.dynamics.com`, or '' standalone. */
  orgUrl: string;
  /** Pinned boundaries keep the preview stable while the user edits. */
  boundaries?: { batch: string; changeset: (i: number) => string };
}

/** Group compiled items back into encodable groups, preserving group order. */
export function toEncodableGroups(items: CompiledItem[]): EncodableGroup[] {
  const groups: EncodableGroup[] = [];
  let last: CompiledItem['group'] | null = null;
  for (const c of items) {
    if (!c.wire) continue;
    if (c.group !== last) groups.push({ atomic: c.group.atomic, parts: [] });
    groups[groups.length - 1].parts.push({ contentId: c.contentId, wire: c.wire });
    last = c.group;
  }
  return groups;
}

function outerHeaders(contentType: string, continueOnError: boolean): string[] {
  return [
    'OData-MaxVersion: 4.0',
    'OData-Version: 4.0',
    'If-None-Match: null',
    'Accept: application/json',
    `Content-Type: ${contentType}`,
    ...(continueOnError ? ['Prefer: odata.continue-on-error'] : []),
  ];
}

const showCrlf = (s: string) => s.replace(/\r\n/g, '\n');

export function generateBatchCode(fmt: BatchCodeFormat, input: BatchCodeInput): string {
  const org = input.orgUrl || '[Organization URI]';
  const encoded = encodeBatch(toEncodableGroups(input.items), input.boundaries);
  switch (fmt) {
    case 'http':
      return showCrlf(
        [
          `POST ${org}${API_BASE}/$batch HTTP/1.1`,
          ...outerHeaders(encoded.contentType, input.continueOnError),
          '',
          encoded.body,
        ].join(CRLF),
      );
    case 'fetch': {
      const lines = encoded.body.split(CRLF);
      if (lines[lines.length - 1] === '') lines.pop();
      const headerLines = outerHeaders(encoded.contentType, input.continueOnError)
        .filter((h) => !h.startsWith('If-None-Match'))
        .map((h) => {
          const i = h.indexOf(':');
          return `    ${JSON.stringify(h.slice(0, i))}: ${JSON.stringify(h.slice(i + 1).trim())},`;
        });
      return [
        '// Dataverse $batch — every line of the body must end in CRLF.',
        `const body = [`,
        ...lines.map((l) => `  ${JSON.stringify(l)},`),
        `].join('\\r\\n') + '\\r\\n';`,
        '',
        `const res = await fetch(${JSON.stringify(`${org}${API_BASE}/$batch`)}, {`,
        `  method: 'POST',`,
        `  headers: {`,
        `    Authorization: \`Bearer \${accessToken}\`,`,
        ...headerLines,
        `  },`,
        `  body,`,
        `});`,
        '',
        '// The response is multipart/mixed too: one part per request, change sets nested.',
        `console.log(res.status, res.headers.get('Content-Type'));`,
        `console.log(await res.text());`,
      ].join('\n');
    }
    case 'pptb': {
      if (input.plan.length === 0) return '// Add a request to see the host calls.';
      const out: string[] = [
        '// What Dataverse Request Studio sends through Power Platform ToolBox (1.2.7+).',
        '// The host writes the multipart body; each change set is its own call.',
      ];
      if (input.plan.length > 1) {
        out.push(
          `// ${input.plan.length} calls, in order.${input.continueOnError ? '' : ' A failed call stops the rest.'}`,
        );
      }
      input.plan.forEach((seg, i) => {
        const reqs = seg.items.filter((c) => c.wire).map(toHostRequest);
        const fn = seg.transaction ? 'executeTransaction' : 'executeBatch';
        const outer =
          !seg.transaction && input.continueOnError
            ? `, undefined, { Prefer: 'odata.continue-on-error' }`
            : '';
        out.push(
          '',
          `const results${i + 1} = await dataverseAPI.${fn}(${JSON.stringify(reqs, null, 2)}${outer});`,
        );
      });
      out.push('', '// Each result: { contentId, status, statusText, headers, body }');
      return out.join('\n').replace(/\\r\\n/g, '\\n');
    }
  }
}
