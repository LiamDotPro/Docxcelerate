/**
 * Conditions, and the branch element that carries one.
 *
 * Nothing here is written by hand. It is what the branch compiler rewrites an
 * `if` into once it has traced the condition back to the request data.
 *
 * @module
 */

import type {
  ComparisonOperator,
  Condition,
  DataReference,
  ValueExpression,
} from "../domain/types.ts";
import { createElement, host, type TemplateElement, type Yield } from "./element.ts";

export const Branch = host<BranchProps, "branch">("branch", "Branch");

export interface BranchProps {
  condition: Condition;
  whenTrue?: () => Yield;
  whenFalse?: () => Yield;
}

/**
 * A decision that survives publishing.
 *
 * Nothing calls this by hand. It is what the branch compiler rewrites an `if`
 * into once it has traced the condition back to the request data, and it exists
 * as a runtime so that the rewrite has somewhere to land and so the behaviour
 * can be tested before the compiler exists.
 *
 * Building a document with real data evaluates the condition and takes one arm,
 * exactly as the `if` would have. Publishing keeps both, each carrying the
 * condition that selects it, and the engine decides per document.
 */
export function branch(
  condition: Condition,
  whenTrue?: () => Yield,
  whenFalse?: () => Yield,
): TemplateElement<"branch"> {
  return createElement("branch", Branch, {
    condition,
    whenTrue,
    whenFalse,
  }) as TemplateElement<"branch">;
}

// Named for the path rather than the reference, because `dataRef` already means
// a value expression in the deriver vocabulary and the two travel together.
/**
 * A path into the data the caller supplied.
 *
 * @param path A dotted path, such as `tenant.name`.
 * @returns The reference.
 */
export function dataPath(path: string): DataReference {
  return { scope: "data", path };
}

/**
 * A path into what the surrounding repeat or component bound.
 *
 * @param path A dotted path, such as `charge.amount`.
 * @returns The reference.
 */
export function ctxPath(path: string): DataReference {
  return { scope: "ctx", path };
}

/**
 * A path into what a deriver wrote.
 *
 * @param path A dotted path, such as `total`.
 * @returns The reference.
 */
export function derivedPath(path: string): DataReference {
  return { scope: "derived", path };
}

/**
 * A condition that holds when the value at a path is truthy.
 *
 * @param ref What to read.
 * @returns The condition.
 */
export function truthy(ref: DataReference): Condition {
  return { type: "truthy", ref };
}

/**
 * A condition comparing two values.
 *
 * @param left The left-hand side.
 * @param operator How the two are measured.
 * @param right The right-hand side.
 * @returns The condition.
 */
export function compare(
  left: ValueExpression,
  operator: ComparisonOperator,
  right: ValueExpression,
): Condition {
  return { type: "compare", operator, left, right };
}

/**
 * A value fixed at build time, for one side of a {@linkcode compare}.
 *
 * @param value The value to carry.
 * @returns The expression.
 */
export function literal(value: string | number | boolean): ValueExpression {
  return { type: "literal", value };
}

/**
 * A value read per document, for one side of a {@linkcode compare}.
 *
 * @param ref What to read.
 * @returns The expression.
 */
export function refValue(ref: DataReference): ValueExpression {
  return { type: "ref", ref };
}

/**
 * A condition that holds when every one of its parts does.
 *
 * @param conditions The parts.
 * @returns The condition.
 */
export function and(...conditions: Condition[]): Condition {
  return { type: "and", conditions };
}

/**
 * A condition that holds when any one of its parts does.
 *
 * @param conditions The parts.
 * @returns The condition.
 */
export function or(...conditions: Condition[]): Condition {
  return { type: "or", conditions };
}
