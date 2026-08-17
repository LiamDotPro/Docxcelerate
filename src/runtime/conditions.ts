import type { ComparisonOperator, Condition, RuntimeState } from "../domain/types.ts";
import { resolveReference, resolveValueExpression } from "./templates.ts";

/**
 * Deciding, per document, what a build left open.
 *
 * @module
 */

/**
 * Evaluates a published condition against the data at hand.
 *
 * @param condition The condition. An absent one holds, so a node with no
 * `when` is always included.
 * @param state The data, context and derived values to test against.
 * @returns Whether the condition holds.
 */
export async function evaluateCondition(
  condition: Condition | undefined,
  state: RuntimeState,
): Promise<boolean> {
  if (!condition) {
    return true;
  }

  if (condition.type === "truthy") {
    return Boolean(await resolveReference(condition.ref, state));
  }

  if (condition.type === "not") {
    return !(await resolveReference(condition.ref, state));
  }

  if (condition.type === "negate") {
    return !(await evaluateCondition(condition.condition, state));
  }

  if (condition.type === "and") {
    for (const nested of condition.conditions) {
      if (!(await evaluateCondition(nested, state))) {
        return false;
      }
    }

    return true;
  }

  if (condition.type === "or") {
    for (const nested of condition.conditions) {
      if (await evaluateCondition(nested, state)) {
        return true;
      }
    }

    return false;
  }

  return compare(
    condition.operator,
    await resolveValueExpression(condition.left, state),
    await resolveValueExpression(condition.right, state),
  );
}

/**
 * Inverts a condition without wrapping it, where an inverse exists.
 *
 * Both arms of a compiled `if` are published, so every branch emits a condition
 * and its opposite. Collapsing `negate` into the operator keeps the published
 * form readable, and keeps the `not` shape reachable for engines that only know
 * the original two.
 */
export function invertCondition(condition: Condition): Condition {
  if (condition.type === "truthy") {
    return { type: "not", ref: condition.ref };
  }

  if (condition.type === "not") {
    return { type: "truthy", ref: condition.ref };
  }

  if (condition.type === "negate") {
    return condition.condition;
  }

  if (condition.type === "compare") {
    return {
      ...condition,
      operator: invertOperator(condition.operator),
    };
  }

  if (condition.type === "and") {
    return { type: "or", conditions: condition.conditions.map(invertCondition) };
  }

  return { type: "and", conditions: condition.conditions.map(invertCondition) };
}

function invertOperator(operator: ComparisonOperator): ComparisonOperator {
  const inverses: Record<ComparisonOperator, ComparisonOperator> = {
    eq: "ne",
    ne: "eq",
    gt: "lte",
    gte: "lt",
    lt: "gte",
    lte: "gt",
  };

  return inverses[operator];
}

function compare(operator: ComparisonOperator, left: unknown, right: unknown): boolean {
  if (operator === "eq") {
    return left === right;
  }

  if (operator === "ne") {
    return left !== right;
  }

  // Ordering only means something between two values of the same comparable
  // kind. Anything else is a question without an answer, so it is false rather
  // than a coerced guess.
  if (typeof left === "number" && typeof right === "number") {
    return orderNumbers(operator, left, right);
  }

  if (typeof left === "string" && typeof right === "string") {
    return orderNumbers(operator, left.localeCompare(right), 0);
  }

  return false;
}

function orderNumbers(operator: ComparisonOperator, left: number, right: number): boolean {
  if (operator === "gt") {
    return left > right;
  }

  if (operator === "gte") {
    return left >= right;
  }

  if (operator === "lt") {
    return left < right;
  }

  return left <= right;
}
