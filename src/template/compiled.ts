/**
 * What compiled control flow calls into.
 *
 * Nothing here is written by hand. The transform rewrites an `if` in a component
 * into a {@linkcode branch} whose condition is built from these, and they exist
 * as a runtime so that the rewrite has somewhere to land and so the behaviour can
 * be tested on its own.
 *
 * Each one hands back a {@linkcode Condition} or a plain boolean, and which it is
 * says everything. A build holding real data settles the test immediately, so the
 * answer is a boolean and the `if` means exactly what it said. A publish build
 * meets a stand-in instead, which knows the path it stands for, so the answer is
 * a condition and the decision travels to the engine to be made per document.
 *
 * The transform supplies the operators, because it can see `!`, `===` and `&&` in
 * the source. The thunks supply the paths, because only a value knows where it
 * came from. Neither has to trace anything back through the component.
 *
 * @module
 */

import type { ComparisonOperator, Condition } from "../domain/types.ts";
import { expr, publishRefOf } from "./publish.ts";

/** A test that is either already settled or left for the engine. */
export type CompiledTest = Condition | boolean;

/**
 * Whether a value is there, as an `if` on it would have asked.
 *
 * @param value The value the condition was written about.
 * @returns The decision, or the condition that makes it per document.
 */
export function __test(value: unknown): CompiledTest {
  const ref = publishRefOf(value);

  return ref ? { type: "truthy", ref } : Boolean(value);
}

/**
 * The opposite of a test.
 *
 * @param test The test to invert.
 * @returns The inverted decision, or the inverted condition.
 */
export function __not(test: CompiledTest): CompiledTest {
  if (typeof test === "boolean") {
    return !test;
  }

  return { type: "negate", condition: test };
}

/**
 * Two values measured against each other, as `===`, `>` or `<=` would have.
 *
 * Both sides go through the same bridge a deriver input does, so a real value
 * travels as a literal and a stand-in as the reference it stands for. A
 * comparison of two real values is still published as a comparison rather than
 * settled, because settling it would bake this build's answer into every
 * document — and the engine measures two literals the same way anyone would.
 *
 * @param left The left-hand side.
 * @param operator How the two are measured.
 * @param right The right-hand side.
 * @returns The condition.
 */
export function __compare(
  left: unknown,
  operator: ComparisonOperator,
  right: unknown,
): CompiledTest {
  return { type: "compare", operator, left: expr(left), right: expr(right) };
}

/**
 * Every test has to hold, as `&&` would have asked.
 *
 * Settled parts are folded away rather than published: a `false` decides the
 * whole thing, and a `true` says nothing the engine needs to be told.
 *
 * @param tests The tests to join.
 * @returns The decision, or the condition that makes it per document.
 */
export function __and(...tests: CompiledTest[]): CompiledTest {
  const open: Condition[] = [];

  for (const test of tests) {
    if (test === false) {
      return false;
    }

    if (test !== true) {
      open.push(test);
    }
  }

  if (open.length === 0) {
    return true;
  }

  return open.length === 1 ? open[0] : { type: "and", conditions: open };
}

/**
 * Any test has to hold, as `||` would have asked.
 *
 * @param tests The tests to join.
 * @returns The decision, or the condition that makes it per document.
 */
export function __or(...tests: CompiledTest[]): CompiledTest {
  const open: Condition[] = [];

  for (const test of tests) {
    if (test === true) {
      return true;
    }

    if (test !== false) {
      open.push(test);
    }
  }

  if (open.length === 0) {
    return false;
  }

  return open.length === 1 ? open[0] : { type: "or", conditions: open };
}
