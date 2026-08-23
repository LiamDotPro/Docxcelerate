/**
 * A deriver as a module: one file, one computation, its own name.
 *
 * A deriver is the part of a document that has to run where the data is. That
 * makes it the part most worth keeping out of a component — a closure inside one
 * node cannot be tested on its own, cannot be shared with the next document, and
 * cannot be reasoned about without reading the document around it. Written as a
 * module it is an ordinary function with an ordinary test beside it.
 *
 * The shape a workspace expects is a folder of them, each with a default export:
 *
 * ```text
 * workspace/
 *   derivers/
 *     invoice-totals.ts
 *     invoice-totals.test.ts
 *     payment-qr.ts
 *   documents/
 *     invoice/
 * ```
 *
 * Nothing about that folder is magic — a deriver is a value, and a document
 * imports the one it uses. The folder is where they go so that two documents can
 * share one, and so the set of them can be packaged and shipped somewhere else
 * to run.
 *
 * @module
 */

import type { DeriverFunction, DeriverDefinition } from "./derivers.ts";

/**
 * A deriver, its name, and what a preview shows instead of running it.
 *
 * @typeParam TInputs What it is called with, in order.
 * @typeParam TResult What it produces.
 */
export interface DeriverModule<TInputs extends readonly unknown[] = readonly unknown[], TResult = unknown>
  extends DeriverDefinition {
  /** The name invocations refer to. */
  readonly name: string;
  /** The computation, taking its inputs as an array the way the engine calls it. */
  readonly run: DeriverFunction;
  /** What a preview shows instead of running this. */
  readonly placeholder?: TResult;
  /** Present only so the types survive; never read. */
  readonly __types?: (inputs: TInputs) => TResult;
}

/** What {@linkcode deriver} takes. */
export interface DefineDeriverOptions<TInputs extends readonly unknown[], TResult> {
  /** The name invocations refer to, and the key its result lands under. */
  name: string;
  /**
   * The computation, written with the arguments it actually takes.
   *
   * Positional rather than an array, because that is what a test calls and what
   * a reader expects. The engine calls every deriver with an array, and the
   * unpacking happens here so nothing else has to think about it.
   */
  run: (...inputs: TInputs) => TResult | Promise<TResult>;
  /**
   * What a preview shows instead of running this.
   *
   * Supplying one is how a deriver says it costs something. A preview is rebuilt
   * every time a file is saved, so a deriver that renders a code, reads a file
   * or asks a service would be paid for on every keystroke. Leave it off for a
   * total or a currency format, which are cheap enough that a preview showing
   * the real figure is worth more than the microsecond.
   */
  placeholder?: TResult;
}

/**
 * Declares a deriver.
 *
 * @typeParam TInputs What it is called with, in order.
 * @typeParam TResult What it produces.
 * @param options Its name, the computation, and what a preview shows instead.
 * @returns The deriver, ready to be the default export of its file.
 *
 * @example
 * ```ts
 * // derivers/invoice-totals.ts
 * import { deriver } from "docxcelerate";
 * import type { Line } from "../documents/invoice/types.ts";
 *
 * export default deriver({
 *   name: "invoiceTotals",
 *   run: (lines: Line[]) => {
 *     const subtotal = lines.reduce((total, line) => total + line.amount, 0);
 *     return { subtotal, vat: subtotal * 0.2, due: subtotal * 1.2 };
 *   },
 * });
 * ```
 */
export function deriver<TInputs extends readonly unknown[], TResult>(
  options: DefineDeriverOptions<TInputs, TResult>,
): DeriverModule<TInputs, TResult> {
  if (!options.name || options.name.trim() === "") {
    throw new Error(
      "A deriver needs a name. It is what a published document calls it by, and what the " +
        "engine looks it up under, so it travels with the document rather than with the file.",
    );
  }

  return {
    name: options.name,
    run: (inputs) => options.run(...(inputs as unknown as TInputs)),
    placeholder: options.placeholder,
  };
}

/**
 * Whether a value is a deriver declared by {@linkcode deriver}.
 *
 * @param value The value to test.
 * @returns `true` when it is one.
 */
export function isDeriverModule(value: unknown): value is DeriverModule {
  return Boolean(
    value &&
      typeof value === "object" &&
      typeof (value as DeriverModule).name === "string" &&
      typeof (value as DeriverModule).run === "function",
  );
}

/**
 * What a deriver produces, for the components that read its result.
 *
 * @typeParam TModule The deriver being read.
 */
export type DeriverResult<TModule> = TModule extends DeriverModule<readonly unknown[], infer R> ? R
  : unknown;

/**
 * What a deriver takes, for the components that call it.
 *
 * @typeParam TModule The deriver being called.
 */
export type DeriverInputs<TModule> = TModule extends DeriverModule<infer I, unknown> ? I
  : readonly unknown[];
