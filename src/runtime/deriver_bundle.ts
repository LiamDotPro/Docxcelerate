import type { DocumentModel, DocumentNode } from "../domain/types.ts";
import {
  createDefaultDeriverRegistry,
  createDeriverRegistry,
  type DeriverDefinition,
  type DeriverDefinitions,
  DeriverRegistry,
  normalizeDeriverDefinitions,
} from "./derivers.ts";

/**
 * Carrying derivers to the engine that runs them.
 *
 * A published document names the derivers it invokes; the functions travel
 * beside it as plain ESM source, because the engine resolving the document is
 * not the process that wrote it and has no other way to get at them.
 *
 * Serializing a function is the part that limits what a deriver may be: it is
 * read back from its own source, so it cannot close over anything.
 *
 * @module
 */
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

  const mod = await import(sourceModuleDataUrl(definitions.source));
  const exported = (mod.default ?? mod.derivers) as DeriverDefinitions | undefined;

  if (!exported) {
    throw new Error("Deriver bundle must export default derivers or a named derivers export.");
  }

  return createDeriverRegistry(exported);
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

