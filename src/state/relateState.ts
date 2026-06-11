// State shapes for the Relate group (Associate / Disassociate).
//
// Both modes work over a navigation property defined on the source entity.
// Cardinality is METADATA-DRIVEN, not picker-driven:
//   • ManyToOne   (lookup column on source) → single-valued nav prop
//   • OneToMany   (source is the "one")     → collection-valued
//   • ManyToMany  (intersect table)         → collection-valued
//
// The URL shape + HTTP verb fall out of the cardinality automatically; the UI
// surfaces this so the user doesn't have to think about it. Per docs:
//
//   • Associate collection-valued: POST /<set>(<id>)/<nav>/$ref     · body: {@odata.id}
//   • Associate single-valued:     PUT  /<set>(<id>)/<nav>/$ref     · body: {@odata.id}
//   • Disassociate collection:     DELETE /<set>(<id>)/<nav>(<rid>)/$ref
//   • Disassociate single:         DELETE /<set>(<id>)/<nav>/$ref
//
// Reference:
//   https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/associate-disassociate-entities-using-web-api

import type { HeaderItem } from '../editors/HeadersEditor';

// ──────────────────────────────────────────────────────────────
// Associate
// ──────────────────────────────────────────────────────────────
//
// Multiple targets are allowed for collection-valued nav props — the builder
// emits one POST request per target. Single-valued nav props accept exactly
// one target.
export interface AssociateState {
  /** Source table logical name (the "From" side of the relationship). */
  table: string;
  /** GUID of the source row. */
  sourceId: string | null;
  /** Selected navigation property name on the source entity (or null until picked). */
  navProperty: string | null;
  /**
   * Target record GUIDs. For collection-valued nav props, 1+ targets queue
   * 1+ POST requests. For single-valued, only the first (and only one) is
   * used in a PUT.
   */
  targets: string[];
  /**
   * Primary-name cache for the picked targets, keyed by GUID. Populated by
   * the AssociateTargetsEditor when the user picks via the live typeahead
   * (the picker hands back the primary name alongside the id). Used purely
   * for display in the selected-targets card and the WriteResultCard
   * narrative — the wire request doesn't include names. Missing entries
   * render as "(name not resolved)" with the GUID still visible.
   */
  targetNames: Record<string, string>;
  /** HTTP headers. */
  headers: HeaderItem[];
  /** Dirty bits for sidebar badges. */
  dirty: Set<string>;
}

// ──────────────────────────────────────────────────────────────
// Disassociate
// ──────────────────────────────────────────────────────────────
//
// Collection-valued: 1+ target ids — each fires its own DELETE call per docs.
// Single-valued:     targetIds ignored (server nulls the lookup; one call only).
export interface DisassociateState {
  /** Source table logical name. */
  table: string;
  /** GUID of the source row. */
  sourceId: string | null;
  /** Selected navigation property name. */
  navProperty: string | null;
  /**
   * Target row GUIDs to disassociate.
   *
   *   • Collection-valued (1:N / N:N): 1+ targets — N sequential DELETEs.
   *   • Single-valued (N:1): ignored — the body's `<nav>@odata.bind: null`
   *     clears the lookup; no target id involved.
   */
  targetIds: string[];
  /**
   * Primary-name cache for the picked targets (collection-valued only),
   * keyed by GUID. Same purpose as AssociateState.targetNames — display
   * + saved-request rehydration.
   */
  targetNames: Record<string, string>;
  /** HTTP headers. */
  headers: HeaderItem[];
  /** Dirty bits for sidebar badges. */
  dirty: Set<string>;
}

// ──────────────────────────────────────────────────────────────
// Helpers
// ──────────────────────────────────────────────────────────────
import type { NavProperty } from '../mock/metadata';

/** True for OneToMany / ManyToMany nav props (collection-valued). */
export const isCollectionValuedNav = (n: NavProperty): boolean =>
  n.cardinality === 'OneToMany' || n.cardinality === 'ManyToMany';

/** True for ManyToOne nav props (single-valued lookup on the source). */
export const isSingleValuedNav = (n: NavProperty): boolean => n.cardinality === 'ManyToOne';

/** Method that the spec dictates for Associate on this cardinality. */
export const associateMethodFor = (n: NavProperty): 'POST' | 'PUT' =>
  isSingleValuedNav(n) ? 'PUT' : 'POST';
