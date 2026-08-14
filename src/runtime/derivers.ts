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

export type DeriverFunction = (
  inputs: unknown[],
  state: RuntimeState,
) => Promise<unknown> | unknown;

export interface DeriverDefinition {
  name: string;
  run: DeriverFunction;
}

export type DeriverDefinitions =
  | Record<string, DeriverFunction>
  | readonly DeriverDefinition[];

export interface DocumentDeriverBundle {
  schemaVersion: "docxcelerate.deriver-bundle/v0";
  format: "esm";
  names: string[];
  source: string;
  entrypoint?: string;
  bundledAt?: string;
}

export interface CreateDeriverBundleOptions {
  names?: Iterable<string>;
  entrypoint?: string;
  bundledAt?: string;
}

export class DeriverRegistry {
  readonly #derivers = new Map<string, DeriverFunction>();

  register(name: string, deriver: DeriverFunction): this {
    this.#derivers.set(name, deriver);
    return this;
  }

  has(name: string): boolean {
    return this.#derivers.has(name);
  }

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

export function defineDeriver(name: string, run: DeriverFunction): DeriverDefinition {
  return { name, run };
}

export function defineDerivers(definitions: DeriverDefinitions): DeriverDefinitions {
  return definitions;
}

export function listDeriverDefinitionNames(definitions: DeriverDefinitions | undefined): string[] {
  return definitions
    ? normalizeDeriverDefinitions(definitions).map((definition) => definition.name)
    : [];
}

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

export function collectDocumentDeriverNames(doc: DocumentModel): string[] {
  const names = new Set<string>();
  collectNodeDeriverNames(doc.nodes, names);
  return [...names].sort((left, right) => left.localeCompare(right));
}

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

export function dataRef(path: string): ValueExpression {
  return ref("data", path);
}

export function ctxRef(path: string): ValueExpression {
  return ref("ctx", path);
}

export function derivedRef(path: string): ValueExpression {
  return ref("derived", path);
}

export function literalValue(value: string | number | boolean): ValueExpression {
  return { type: "literal", value };
}

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

    if (node.kind === "section") {
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
