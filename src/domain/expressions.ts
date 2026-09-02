/**
 * Saying what a value is, and what has to be true, without knowing either yet.
 *
 * A document is built before the data it will be generated against exists, so
 * the decisions it cannot make are published rather than taken: a condition to
 * test, an expression naming where a value will be found, an invocation naming
 * a computation to run. An engine settles all three later.
 *
 * @module
 */

/**
 * What a prompt is for, which is how an engine decides where to put it.
 *
 * `general` asks for something, `info` supplies facts to write from, `negative`
 * rules something out, `example` shows the shape a good answer takes, and
 * `system` sets the standing instructions.
 */
export type PromptKind = "example" | "general" | "info" | "negative" | "system";

/**
 * Which bag of values a {@linkcode DataReference} reads from.
 *
 * `data` is what the caller supplied, `ctx` is what the surrounding repeat or
 * component bound, and `derived` is what a {@linkcode DeriverInvocation} wrote.
 */
export type ReferenceScope = "data" | "ctx" | "derived";

/** A pointer to a value that only exists once a document is being written. */
export interface DataReference {
  /** Which bag of values to read from. */
  scope: ReferenceScope;
  /** A dotted path into that bag, such as `tenant.name`. */
  path: string;
}

/** How the two sides of a `compare` {@linkcode Condition} are measured. */
export type ComparisonOperator = "eq" | "ne" | "gt" | "gte" | "lt" | "lte";

/**
 * A decision the engine makes per document, rather than one the build makes once.
 *
 * `truthy` and `not` came first and are still emitted for the shapes they cover,
 * so an engine that predates the richer forms keeps understanding the common
 * case. The compiler that turns an `if` in a component into one of these picks
 * the narrowest form that fits.
 */
export type Condition =
  | { type: "truthy"; ref: DataReference }
  | { type: "not"; ref: DataReference }
  | {
    type: "compare";
    operator: ComparisonOperator;
    left: ValueExpression;
    right: ValueExpression;
  }
  | { type: "and"; conditions: Condition[] }
  | { type: "or"; conditions: Condition[] }
  | { type: "negate"; condition: Condition };

/**
 * Either a value written into the document at build time or a pointer to one
 * that will not exist until the document is written.
 */
export type ValueExpression =
  | { type: "literal"; value: string | number | boolean }
  | { type: "ref"; ref: DataReference };

/** One instruction attached to a node, for an engine that writes its content. */
export interface PromptSpec {
  /** What the prompt is for. */
  kind: PromptKind;
  /** The prompt itself. */
  text: string;
}

/**
 * A named computation to run before a node is written, and where to put the
 * result.
 *
 * Deriving happens on the engine because the inputs do too. The result lands in
 * the `derived` scope, so anything downstream reads it the same way it reads
 * caller data.
 */
export interface DeriverInvocation {
  /** The key the result is written to under the `derived` scope. */
  output: string;
  /** Which registered deriver to run. */
  name: string;
  /** The arguments to call it with, in order. */
  inputs: ValueExpression[];
}
