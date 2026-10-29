// What a write's confirmation says about its target, and when one is asked
// for (#241). Pure, so the rules and the wording are unit-tested.
//
// PPTB Marketplace policy for tools that change data: confirmations name the
// target environment and the scope (which table, how many records).
// https://docs.powerplatformtoolbox.com/policies/marketplace
// The rules match FetchXML Studio's targetScopeRules (its #69).

/** More records than this need the confirmation word typed. */
export const TYPE_TO_CONFIRM_RECORD_THRESHOLD = 100;

/** What the request applies to. */
export type WriteScope =
  /** These records (the request's targets, or rows sent). */
  | { kind: 'records'; count: number }
  /** Anything else, already worded (an unbound action, a batch's requests). */
  | { kind: 'text'; text: string };

/**
 * What a mode tells the Execute button about the request it's about to send.
 * `null` from a mode means the request doesn't change data (or the mode asks
 * on its own, like Delete's typed confirmation).
 */
export interface WriteConfirmSpec {
  /** The operation, as the user knows it: "Merge Records", "Execute Action CanBeReferenced". */
  operation: string;
  /** Table logical name, when the request targets one. */
  table?: string;
  scope: WriteScope;
  /**
   * Irreversible or high-impact (Merge, actions, workflows, bulk and batch
   * writes, multi-target relates): always confirmed. Otherwise the write is
   * confirmed only on a Production connection with safe mode on.
   */
  highImpact: boolean;
  /**
   * What the user types when a typed confirmation is required. Defaults to
   * the table logical name, then the operation.
   */
  confirmWord?: string;
}

/** PPTB connection environment types: "Dev" | "Test" | "UAT" | "Production". */
export const isProductionEnvironment = (environment: string | undefined): boolean =>
  environment?.trim().toLowerCase() === 'production';

/** Whether Execute asks before sending this request. */
export function needsWriteConfirmation(
  spec: WriteConfirmSpec | null | undefined,
  environment: string | undefined,
  productionSafeMode: boolean,
): boolean {
  if (!spec) return false;
  if (spec.highImpact) return true;
  return productionSafeMode && isProductionEnvironment(environment);
}

/**
 * Whether the confirmation word must be typed before the button enables: on
 * Production with safe mode on, and for more than
 * TYPE_TO_CONFIRM_RECORD_THRESHOLD records anywhere.
 */
export function requiresTypedConfirmation(
  spec: WriteConfirmSpec,
  environment: string | undefined,
  productionSafeMode: boolean,
): boolean {
  if (productionSafeMode && isProductionEnvironment(environment)) return true;
  return spec.scope.kind === 'records' && spec.scope.count > TYPE_TO_CONFIRM_RECORD_THRESHOLD;
}

/** The word to type: the given one, else the table, else the operation. */
export const confirmWordOf = (spec: WriteConfirmSpec): string =>
  (spec.confirmWord ?? spec.table ?? spec.operation).trim();

/** The scope as one line. */
export function describeWriteScope(scope: WriteScope): string {
  if (scope.kind === 'text') return scope.text;
  return scope.count === 1 ? '1 record' : `${scope.count.toLocaleString()} records`;
}

/** Whether the typed text confirms: the expected word, ignoring case and surrounding spaces. */
export const isTypedConfirmation = (typed: string, expected: string): boolean =>
  expected.trim() !== '' && typed.trim().toLowerCase() === expected.trim().toLowerCase();
