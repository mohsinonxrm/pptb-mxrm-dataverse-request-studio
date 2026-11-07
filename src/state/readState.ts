// Shared state shape for the Read group. Each mode picks the subset it needs.
import type { FilterGroup } from '../editors/filter/filterTree';
import type { OrderbySpec } from '../editors/OrderbyEditor';
import type { ExpandSpec } from '../editors/ExpandEditor';
import type { ApplySpec } from '../editors/ApplyEditor';
import type { PreferSpec } from '../editors/PreferEditor';
import type { HeaderItem } from '../editors/HeadersEditor';

export interface RetrieveMultipleState {
  table: string;
  select: string[];
  filter: FilterGroup;
  orderby: OrderbySpec[];
  top: number | null;
  countOn: boolean;
  /**
   * Count only (#272): just the number of rows matching $filter (up to
   * 5,000), read from `@odata.count` on a one-row request. The other clauses
   * aren't sent. Optional, so saved requests from before it read as off.
   */
  countOnly?: boolean;
  expand: ExpandSpec[];
  apply: ApplySpec;
  prefer: PreferSpec;
  headers: HeaderItem[];
  /** Tracks which clauses were edited since the last execute */
  dirty: Set<string>;
}

export interface RetrieveSingleState {
  table: string;
  recordId: string | null;
  select: string[];
  expand: ExpandSpec[];
  prefer: PreferSpec;
  headers: HeaderItem[];
  dirty: Set<string>;
}

export interface RetrieveNextLinkState {
  url: string;
  prefer: PreferSpec;
  headers: HeaderItem[];
  dirty: Set<string>;
}

export interface PredefinedQueryState {
  table: string;
  queryId: string | null;
  queryType: 'savedQuery' | 'userQuery';
  top: number | null;
  prefer: PreferSpec;
  headers: HeaderItem[];
  dirty: Set<string>;
  /**
   * Apply the query to a related collection of one row instead of the
   * table's entity set (#5): `/<sourceSet>(<id>)/<nav>?savedQuery=<id>`.
   * Null / absent (entries saved before #5) = the entity set.
   */
  related?: PredefinedRelatedTarget | null;
}

/** The row + collection-valued nav a predefined query runs against (#5). */
export interface PredefinedRelatedTarget {
  /** Table of the source row (logical name), e.g. `account`. */
  sourceTable: string;
  /** Relationship schema name — identifies the pair of nav properties. */
  relationshipName: string;
  /** Collection-valued nav on the source table, e.g. `opportunity_parent_account`. */
  navProperty: string;
  /** Source row id. */
  recordId: string | null;
  /** Source row primary name, for display. */
  recordName?: string;
}

// Recent runs (in-memory, optional persistence later)
export interface RecentRun {
  id: string;
  modeId: string;
  url: string;
  method: string;
  ts: number;
  status: number;
  /** Replaces the "N rows" text — e.g. a batch's "3 requests · 1 failed". */
  summary?: string;
  ms: number;
  rowCount?: number;
  /**
   * Serialized mode state at execute time (same shape the Saved library
   * stores). Lets the sidebar restore the run on click (#2). Absent for
   * modes without a serializer (e.g. Retrieve NextLink).
   */
  snapshot?: unknown;
  /**
   * The kept copy of this run had its impersonation caller id removed
   * (#226), so restoring it runs as the signed-in user unless it's set again.
   */
  impersonationCleared?: boolean;
}
