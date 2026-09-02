import type { DeriverInvocation, RuntimeState } from "../domain/types.ts";
import { setPath } from "./object_path.ts";
import { resolveValueExpression } from "./templates.ts";

/**
 * The derivers a build can run, and running them.
 *
 * A deriver runs where the data is, not where the document was written, so a
 * document carries invocations — a name, its arguments and where to put the
 * answer — rather than functions. This module is the registry those names are
 * looked up in. Getting the functions to an engine is `deriver_bundle.ts`, and
 * writing the arguments is `deriver_refs.ts`.
 *
 * @module
 */

/**
 * A computation the engine can run while writing a document.
 *
 * @param inputs The invocation's arguments, already resolved against the data.
 * @param state Everything reachable while this document is being written.
 * @returns The value to store under the invocation's output key.
 */
export type DeriverFunction = (
  inputs: unknown[],
  state: RuntimeState,
) => Promise<unknown> | unknown;

/** A deriver paired with the name documents call it by. */
export interface DeriverDefinition {
  /** The name invocations refer to. */
  name: string;
  /** The computation itself. */
  run: DeriverFunction;
  /**
   * What a preview shows instead of running this.
   *
   * Supplying one is how a deriver says it costs something. A preview is
   * rebuilt every time a file is saved, so anything it waits for is time
   * somebody spends watching a document fail to appear — a code to render, a
   * file to read, a service to ask. Those stand in while a document is being
   * written, and run for real when one is.
   *
   * A deriver without one is cheap by construction — a total, a currency, a
   * date — and runs everywhere, because a preview showing the real figure is
   * worth more than the microsecond it took.
   */
  placeholder?: unknown;
}

/** A set of derivers, written either as an object or as a list of definitions. */
export type DeriverDefinitions =
  | Record<string, DeriverFunction>
  | readonly DeriverDefinition[];

/** The derivers available while a document is being written, by name. */
export class DeriverRegistry {
  readonly #derivers = new Map<string, DeriverFunction>();
  readonly #placeholders = new Map<string, unknown>();

  /**
   * Adds a deriver, replacing any already registered under the name.
   *
   * @param name The name invocations refer to.
   * @param deriver The computation to run.
   * @returns This registry, so registrations can be chained.
   */
  register(name: string, deriver: DeriverFunction, placeholder?: unknown): this {
    this.#derivers.set(name, deriver);

    if (placeholder !== undefined) {
      this.#placeholders.set(name, placeholder);
    }

    return this;
  }

  /**
   * What a preview shows instead of running one of these.
   *
   * @param name The deriver to ask about.
   * @returns The stand-in, or `undefined` when the deriver is cheap enough to run.
   */
  placeholderFor(name: string): unknown {
    return this.#placeholders.get(name);
  }

  /**
   * Whether a preview should stand in for a deriver rather than run it.
   *
   * @param name The deriver to ask about.
   * @returns `true` when it declared a stand-in.
   */
  standsInForPreview(name: string): boolean {
    return this.#placeholders.has(name);
  }

  /**
   * Whether a deriver is registered under the name.
   *
   * @param name The name to look for.
   * @returns `true` when something is registered under it.
   */
  has(name: string): boolean {
    return this.#derivers.has(name);
  }

  /**
   * Writes a deriver's stand-in where its result would have gone.
   *
   * The shape is the same as the real thing, so anything downstream reads it
   * the same way — a preview differs from a document in what the value says,
   * never in whether it is there.
   *
   * @param invocation Which deriver to stand in for, and where to put it.
   * @param state The state to write into.
   * @returns The stand-in that was written.
   */
  standIn(invocation: DeriverInvocation, state: RuntimeState): unknown {
    const placeholder = this.#placeholders.get(invocation.name);

    setPath(state.derived, invocation.output, placeholder);

    return placeholder;
  }

  /**
   * Resolves an invocation's inputs, runs its deriver, and writes the result
   * into the `derived` scope.
   *
   * @param invocation Which deriver to run, with what, and where to put it.
   * @param state The state to resolve against and write into.
   * @returns The value the deriver produced.
   * @throws If nothing is registered under the invocation's name.
   */
  async run(invocation: DeriverInvocation, state: RuntimeState): Promise<unknown> {
    const deriver = this.#derivers.get(invocation.name);
    if (!deriver) {
      throw new Error(`Unknown deriver: ${invocation.name}`);
    }

    const inputs = await Promise.all(
      invocation.inputs.map((input) => resolveValueExpression(input, state)),
    );
    const value = await deriver(inputs, state);
    setPath(state.derived, invocation.output, value);
    return value;
  }
}

/**
 * A registry holding the derivers every document can assume: `sum`, `join`
 * and `count`.
 *
 * @returns A fresh registry, safe to register over.
 */
export function createDefaultDeriverRegistry(): DeriverRegistry {
  return new DeriverRegistry()
    .register(
      "sum",
      (inputs) => inputs.reduce<number>((total, input) => total + numberOrZero(input), 0),
    )
    .register("join", (inputs) =>
      inputs
        .filter((input) => input !== undefined && input !== null && input !== "")
        .map(String)
        .join(" "))
    .register("count", (inputs) => {
      const [value] = inputs;
      return Array.isArray(value) ? value.length : inputs.length;
    });
}

/**
 * The built-in derivers, with the caller's own registered over them.
 *
 * @param definitions Derivers to add. A registry is passed straight back.
 * @returns The registry to run a document against.
 */
export function createDeriverRegistry(
  definitions?: DeriverDefinitions | DeriverRegistry,
): DeriverRegistry {
  if (definitions instanceof DeriverRegistry) {
    return definitions;
  }

  const registry = createDefaultDeriverRegistry();

  if (definitions) {
    registerDeriverDefinitions(registry, definitions);
  }

  return registry;
}

/**
 * The names a set of definitions provides.
 *
 * @param definitions The derivers to read, in either form.
 * @returns Their names, or an empty array when there are none.
 */
export function listDeriverDefinitionNames(definitions: DeriverDefinitions | undefined): string[] {
  return definitions
    ? normalizeDeriverDefinitions(definitions).map((definition) => definition.name)
    : [];
}

/**
 * Runs a node's invocations in order, so a later one can read what an earlier
 * one wrote.
 *
 * @param invocations The node's derivers, if it has any.
 * @param state The state to resolve against and write into.
 * @param registry Where the derivers are looked up.
 * @param previewing Whether a deriver that declared a stand-in should use it.
 * @throws If an invocation names a deriver the registry does not hold.
 */
async function runDerivers(
  invocations: DeriverInvocation[] | undefined,
  state: RuntimeState,
  registry: DeriverRegistry,
  previewing = false,
): Promise<void> {
  for (const invocation of invocations ?? []) {
    if (previewing && registry.standsInForPreview(invocation.name)) {
      registry.standIn(invocation, state);
      continue;
    }

    await registry.run(invocation, state);
  }
}

export { runDerivers };

function registerDeriverDefinitions(
  registry: DeriverRegistry,
  definitions: DeriverDefinitions | DeriverRegistry,
): void {
  if (definitions instanceof DeriverRegistry) {
    return;
  }

  for (const definition of normalizeDeriverDefinitions(definitions)) {
    registry.register(definition.name, definition.run, definition.placeholder);
  }
}

/**
 * The definitions in either form, as a list.
 *
 * A project may write its derivers as an object keyed by name or as an array of
 * definitions; everything downstream wants the array.
 *
 * @param definitions The derivers to read, in either form.
 * @returns One definition per deriver.
 */
export function normalizeDeriverDefinitions(definitions: DeriverDefinitions): DeriverDefinition[] {
  if (Array.isArray(definitions)) {
    return definitions.map((definition) => ({
      name: definition.name,
      run: definition.run,
      placeholder: definition.placeholder,
    }));
  }

  return Object.entries(definitions).map(([name, run]) => ({ name, run }));
}

function numberOrZero(value: unknown): number {
  const numericValue = Number(value);
  return Number.isFinite(numericValue) ? numericValue : 0;
}
