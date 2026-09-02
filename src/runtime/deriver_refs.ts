import type { DataReference, DeriverInvocation, ValueExpression } from "../domain/types.ts";
import { expr } from "../template/publish.ts";

/**
 * Pointing at a value without having it yet.
 *
 * A deriver invocation is written where the document is composed and run where
 * the data is, so its arguments cannot be values — they are expressions naming
 * where a value will be found. The same expressions are what a condition
 * compares, which is why they live apart from the registry that runs them.
 *
 * @module
 */
/**
 * Builds the invocation a node carries: which deriver to run, with what, and
 * where the answer goes.
 *
 * @param name The registered deriver to run.
 * @param options Where to write the result, and what to pass in.
 * @returns The invocation to attach to a node.
 *
 * @example
 * ```ts
 * derive("sum", { output: "total", inputs: [dataRef("charges.rent")] });
 * ```
 */
export function derive(
  name: string,
  options: { output: string; inputs?: readonly unknown[] },
): DeriverInvocation {
  return {
    name,
    output: options.output,
    inputs: (options.inputs ?? []).map(toValueExpression),
  };
}

/**
 * An input as written: either an expression built by hand, or the value itself.
 *
 * Passing the value is the shorter road and the only one that works inside a
 * `.map()`, where what the component holds is a real entry on one build and a
 * stand-in on the other.
 */
function toValueExpression(input: unknown): ValueExpression {
  if (isValueExpression(input)) {
    return input;
  }

  return expr(input);
}

function isValueExpression(input: unknown): input is ValueExpression {
  if (typeof input !== "object" || input === null) {
    return false;
  }

  const type = (input as { type?: unknown }).type;

  return type === "literal" || type === "ref";
}

/**
 * A pointer into the data the caller supplied.
 *
 * @param path A dotted path, such as `tenant.name`.
 * @returns The expression to use as a deriver input or comparison side.
 */
export function dataRef(path: string): ValueExpression {
  return ref("data", path);
}

/**
 * A pointer into what the surrounding repeat or component bound.
 *
 * @param path A dotted path, such as `charge.amount`.
 * @returns The expression to use as a deriver input or comparison side.
 */
export function ctxRef(path: string): ValueExpression {
  return ref("ctx", path);
}

/**
 * A pointer into what an earlier deriver wrote.
 *
 * @param path A dotted path, such as `total`.
 * @returns The expression to use as a deriver input or comparison side.
 */
export function derivedRef(path: string): ValueExpression {
  return ref("derived", path);
}

/**
 * A value fixed at build time.
 *
 * @param value The value to carry.
 * @returns The expression to use as a deriver input or comparison side.
 */
export function literalValue(value: string | number | boolean): ValueExpression {
  return { type: "literal", value };
}

function ref(scope: DataReference["scope"], path: string): ValueExpression {
  return {
    type: "ref",
    ref: {
      scope,
      path,
    },
  };
}

