import type {
  DataReference,
  DeriverInvocation,
  DocumentModel,
  DocumentNode,
  RuntimeState,
  ValueExpression,
} from "../domain/types.ts";
import { setPath } from "./object_path.ts";
import { resolveValueExpression } from "./templates.ts";

/**
 * Computations a document defers to the engine, and the plumbing that carries
 * them there.
 *
 * A deriver runs where the data is, not where the document was written. So the
 * document publishes an invocation — a name, its arguments and where to put the
 * answer — and the function itself travels separately, as a
 * {@linkcode DocumentDeriverBundle} of plain ESM source.
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
}

/** A set of derivers, written either as an object or as a list of definitions. */
export type DeriverDefinitions =
  | Record<string, DeriverFunction>
  | readonly DeriverDefinition[];

/**
 * Derivers serialized to source, so an engine can load the functions a document
 * needs without the document project being installed anywhere near it.
 */
export interface DocumentDeriverBundle {
  /** The bundle version, so a reader knows what it is looking at. */
  schemaVersion: "docxcelerate.deriver-bundle/v0";
  /** How `source` is written. */
  format: "esm";
  /** The names the bundle defines. */
  names: string[];
  /** An ES module exporting the derivers, both by default and as `derivers`. */
  source: string;
  /** The document project the bundle was built from. */
  entrypoint?: string;
  /** When the bundle was built, as an ISO 8601 string. */
  bundledAt?: string;
}

/** What {@linkcode createDeriverBundle} takes beyond the definitions themselves. */
export interface CreateDeriverBundleOptions {
  /** Bundle only these derivers. Everything is bundled when absent. */
  names?: Iterable<string>;
  /** The document project being bundled, recorded on the bundle. */
  entrypoint?: string;
  /** The timestamp to record, as an ISO 8601 string. */
  bundledAt?: string;
}

/** The derivers available while a document is being written, by name. */
export class DeriverRegistry {
  readonly #derivers = new Map<string, DeriverFunction>();

  /**
   * Adds a deriver, replacing any already registered under the name.
   *
   * @param name The name invocations refer to.
   * @param deriver The computation to run.
   * @returns This registry, so registrations can be chained.
   */
  register(name: string, deriver: DeriverFunction): this {
    this.#derivers.set(name, deriver);
    return this;
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
 * The same as {@linkcode createDeriverRegistry}, but also accepting a
 * {@linkcode DocumentDeriverBundle} — which has to be imported before its
 * derivers exist, hence the promise.
 *
 * @param definitions Derivers, a registry, or a bundle to load.
 * @returns The registry to run a document against.
 * @throws If a bundle's source exports no derivers.
 */
export async function createDeriverRegistryFromBundle(
  definitions?: DeriverDefinitions | DeriverRegistry | DocumentDeriverBundle,
): Promise<DeriverRegistry> {
  if (!definitions) {
    return createDefaultDeriverRegistry();
  }

  if (definitions instanceof DeriverRegistry || !isDocumentDeriverBundle(definitions)) {
    return createDeriverRegistry(definitions);
  }

  const registry = createDefaultDeriverRegistry();
  const mod = await import(sourceModuleDataUrl(definitions.source));
  const exported = (mod.default ?? mod.derivers) as DeriverDefinitions | undefined;

  if (!exported) {
    throw new Error("Deriver bundle must export default derivers or a named derivers export.");
  }

  registerDeriverDefinitions(registry, exported);
  return registry;
}

/**
 * Names a deriver, for the list form of {@linkcode DeriverDefinitions}.
 *
 * @param name The name invocations refer to.
 * @param run The computation to run.
 * @returns The definition.
 */
export function defineDeriver(name: string, run: DeriverFunction): DeriverDefinition {
  return { name, run };
}

/**
 * Types a project's derivers without changing them, so mistakes surface where
 * they are written rather than where they are used.
 *
 * @param definitions The project's derivers.
 * @returns The same value.
 *
 * @example
 * ```ts
 * export const derivers = defineDerivers({
 *   total: (inputs) => inputs.reduce<number>((sum, n) => sum + Number(n), 0),
 * });
 * ```
 */
export function defineDerivers(definitions: DeriverDefinitions): DeriverDefinitions {
  return definitions;
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
 * Serializes derivers to ESM source an engine can import.
 *
 * The functions are read back with `Function.prototype.toString`, so a deriver
 * has to stand on its own: anything it closes over will not be there when the
 * bundle runs.
 *
 * @param definitions The derivers to bundle.
 * @param options Which of them to include, and what to record alongside.
 * @returns The bundle, or `undefined` when nothing was selected.
 * @throws If a selected deriver is native code and cannot be read back.
 */
export function createDeriverBundle(
  definitions: DeriverDefinitions | undefined,
  options: CreateDeriverBundleOptions = {},
): DocumentDeriverBundle | undefined {
  if (!definitions) {
    return undefined;
  }

  const requestedNames = options.names ? new Set(options.names) : undefined;
  const selected = normalizeDeriverDefinitions(definitions).filter((definition) =>
    !requestedNames || requestedNames.has(definition.name)
  );

  if (selected.length === 0) {
    return undefined;
  }

  return {
    schemaVersion: "docxcelerate.deriver-bundle/v0",
    format: "esm",
    names: selected.map((definition) => definition.name),
    source: deriverModuleSource(selected),
    entrypoint: options.entrypoint,
    bundledAt: options.bundledAt,
  };
}

/**
 * Every deriver a built document invokes, found by walking it — including the
 * nodes inside sections and repeats.
 *
 * @param doc The built document.
 * @returns The names, sorted and deduplicated.
 */
export function collectDocumentDeriverNames(doc: DocumentModel): string[] {
  const names = new Set<string>();
  collectNodeDeriverNames(doc.nodes, names);
  return [...names].sort((left, right) => left.localeCompare(right));
}

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
  options: { output: string; inputs?: ValueExpression[] },
): DeriverInvocation {
  return {
    name,
    output: options.output,
    inputs: options.inputs ?? [],
  };
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

/**
 * Runs a node's invocations in order, so a later one can read what an earlier
 * one wrote.
 *
 * @param invocations The node's derivers, if it has any.
 * @param state The state to resolve against and write into.
 * @param registry Where the derivers are looked up.
 * @throws If an invocation names a deriver the registry does not hold.
 */
async function runDerivers(
  invocations: DeriverInvocation[] | undefined,
  state: RuntimeState,
  registry: DeriverRegistry,
): Promise<void> {
  for (const invocation of invocations ?? []) {
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
    registry.register(definition.name, definition.run);
  }
}

function normalizeDeriverDefinitions(definitions: DeriverDefinitions): DeriverDefinition[] {
  if (Array.isArray(definitions)) {
    return definitions.map((definition) => ({
      name: definition.name,
      run: definition.run,
    }));
  }

  return Object.entries(definitions).map(([name, run]) => ({ name, run }));
}

function deriverModuleSource(definitions: DeriverDefinition[]): string {
  const entries = definitions.map((definition) =>
    `  ${JSON.stringify(definition.name)}: ${serializableFunctionSource(definition)},`
  );

  return [
    "const derivers = {",
    ...entries,
    "};",
    "export { derivers };",
    "export default derivers;",
    "",
  ].join("\n");
}

function serializableFunctionSource(definition: DeriverDefinition): string {
  const source = definition.run.toString();

  if (source.includes("[native code]")) {
    throw new Error(`Deriver "${definition.name}" cannot be bundled from native code.`);
  }

  if (
    /^(async\s+)?function\b/.test(source) ||
    /^(async\s+)?\(?[\w\s,[\]{}.:=]*\)?\s*=>/.test(source)
  ) {
    return source;
  }

  if (/^async\s+[A-Za-z_$][\w$]*\s*\(/.test(source)) {
    return source.replace(/^async\s+([A-Za-z_$][\w$]*)\s*\(/, "async function $1(");
  }

  if (/^[A-Za-z_$][\w$]*\s*\(/.test(source)) {
    return `function ${source}`;
  }

  return source;
}

function collectNodeDeriverNames(nodes: DocumentNode[], names: Set<string>): void {
  for (const node of nodes) {
    for (const invocation of node.derivers ?? []) {
      names.add(invocation.name);
    }

    // Both kinds that hold children. Missing one here is not a missing name in
    // a list — it is a deriver left out of the published bundle, and an engine
    // that fails on the document rather than on the build.
    if (node.kind === "section" || node.kind === "repeat") {
      collectNodeDeriverNames(node.children, names);
    }
  }
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

function isDocumentDeriverBundle(value: unknown): value is DocumentDeriverBundle {
  return Boolean(
    value &&
      typeof value === "object" &&
      "schemaVersion" in value &&
      value.schemaVersion === "docxcelerate.deriver-bundle/v0" &&
      "source" in value &&
      typeof value.source === "string",
  );
}

function sourceModuleDataUrl(source: string): string {
  return `data:text/javascript;charset=utf-8,${encodeURIComponent(source)}`;
}

function numberOrZero(value: unknown): number {
  const numericValue = Number(value);
  return Number.isFinite(numericValue) ? numericValue : 0;
}
