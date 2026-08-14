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
export function dataPath(path: string): DataReference {
  return { scope: "data", path };
}

export function ctxPath(path: string): DataReference {
  return { scope: "ctx", path };
}

export function derivedPath(path: string): DataReference {
  return { scope: "derived", path };
}

export function truthy(ref: DataReference): Condition {
  return { type: "truthy", ref };
}

export function compare(
  left: ValueExpression,
  operator: ComparisonOperator,
  right: ValueExpression,
): Condition {
  return { type: "compare", operator, left, right };
}

export function literal(value: string | number | boolean): ValueExpression {
  return { type: "literal", value };
}

export function refValue(ref: DataReference): ValueExpression {
  return { type: "ref", ref };
}

export function and(...conditions: Condition[]): Condition {
  return { type: "and", conditions };
}

export function or(...conditions: Condition[]): Condition {
  return { type: "or", conditions };
}
