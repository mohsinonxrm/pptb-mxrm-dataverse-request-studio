// Who published a Custom API or custom process action (#248).
//
// A typical Dynamics org has well over a thousand Custom APIs and hundreds of
// custom process actions, nearly all from Microsoft solutions, so the Execute
// Custom API / Custom Action pickers group them: your org's and ISVs' first,
// Microsoft's below. Nothing is hidden.
//
// The publisher comes from the row's `solutionid` (both `customapi` and
// `workflow` carry it): the solution that owns the component's current
// layer, and that solution's publisher. A component with an unmanaged layer
// points at the `Active` solution instead, which says nothing about who made
// it, so the message name's customization prefix picks the publisher there.
// Every custom message carries its publisher's prefix (`contoso_DoThing`,
// `new_MyAction`); one without a prefix is Microsoft's.

/** Publisher shown on a Custom API / Custom Action in the picker. */
export interface MessagePublisher {
  /** The publisher's friendly name (or the bare prefix when no publisher has it). */
  name: string;
  /** Listed under "Microsoft" rather than "Your org & ISVs". */
  microsoft: boolean;
}

interface PublisherRow {
  uniquename?: string;
  friendlyname?: string;
  customizationprefix?: string;
}

/** A `solutions` row with its publisher expanded. */
export interface SolutionPublisherRow {
  solutionid?: string;
  uniquename?: string;
  publisherid?: PublisherRow | null;
}

export interface PublisherIndex {
  /** solutionid (lower-case) → the solution's unique name and publisher. */
  solutions: Map<string, { uniquename: string; publisher: MessagePublisher }>;
  /** customization prefix (lower-case) → publisher. */
  byPrefix: Map<string, MessagePublisher>;
}

/** The query that feeds `buildPublisherIndex`. */
export const SOLUTION_PUBLISHERS_QUERY =
  'solutions?$select=solutionid,uniquename' +
  '&$expand=publisherid($select=uniquename,friendlyname,customizationprefix)';

// Microsoft's publishers seen across Dynamics orgs. Unknown ones fall to
// `isMicrosoftPublisher`'s name and prefix checks.
const MICROSOFT_PUBLISHERS = new Set([
  'microsoftcorporation',
  'microsoftfirstparty',
  'microsoftdynamics',
  'microsoftdynamicslabs',
  'dynamics365customerengagement',
  'dynamicsmkt',
  'microsoftformspro',
  'qmsanchor',
  'adxstudio',
]);

// Microsoft's customization prefixes. A customer's or ISV's publisher can
// start with "ms" too (Morgan Stanley, MSC), so a publisher row is matched
// against this list rather than `ms*`.
const MICROSOFT_PREFIXES = new Set([
  'msdyn',
  'msdynce',
  'msdynmkt',
  'msfp',
  'msft',
  'mspp',
  'msa',
  'mscrm',
  'adx',
]);

/** True for a Microsoft publisher: a known one, or one that looks like it. */
export function isMicrosoftPublisher(p: PublisherRow): boolean {
  const unique = (p.uniquename ?? '').toLowerCase();
  if (MICROSOFT_PUBLISHERS.has(unique)) return true;
  if (unique.startsWith('microsoft')) return true;
  if (/^(microsoft|dynamics 365)\b/i.test(p.friendlyname ?? '')) return true;
  return MICROSOFT_PREFIXES.has((p.customizationprefix ?? '').toLowerCase());
}

function toPublisher(p: PublisherRow): MessagePublisher {
  return {
    name: p.friendlyname || p.uniquename || 'Unknown publisher',
    microsoft: isMicrosoftPublisher(p),
  };
}

export function buildPublisherIndex(rows: SolutionPublisherRow[]): PublisherIndex {
  const solutions: PublisherIndex['solutions'] = new Map();
  const byPrefix: PublisherIndex['byPrefix'] = new Map();
  for (const row of rows) {
    if (!row.solutionid || !row.publisherid) continue;
    const publisher = toPublisher(row.publisherid);
    solutions.set(row.solutionid.toLowerCase(), { uniquename: row.uniquename ?? '', publisher });
    const prefix = row.publisherid.customizationprefix?.toLowerCase();
    // Two publishers can share a prefix (Microsoft's do: msdyn); the first
    // one wins.
    if (prefix && !byPrefix.has(prefix)) byPrefix.set(prefix, publisher);
  }
  return { solutions, byPrefix };
}

// Solutions that hold unmanaged layers, not a publisher's work.
const UNMANAGED_LAYER_SOLUTIONS = new Set(['active', 'default']);

/** The publisher of message `name`, owned by solution `solutionId`. */
export function resolvePublisher(
  name: string,
  solutionId: string | undefined,
  index: PublisherIndex,
): MessagePublisher {
  const solution = solutionId ? index.solutions.get(solutionId.toLowerCase()) : undefined;
  if (solution && !UNMANAGED_LAYER_SOLUTIONS.has(solution.uniquename.toLowerCase())) {
    return solution.publisher;
  }
  const prefix = /^([a-z][a-z0-9]*)_/i.exec(name)?.[1];
  if (!prefix) return { name: 'Microsoft', microsoft: true };
  // No publisher has this prefix: only the name is left to go on, so any
  // ms* prefix reads as Microsoft's, like the source heuristic.
  return (
    index.byPrefix.get(prefix.toLowerCase()) ?? {
      name: prefix,
      microsoft: /^(ms[a-z0-9]*|adx)$/i.test(prefix),
    }
  );
}
