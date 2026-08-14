import { buildDocument, type ComponentRuntimeOptions } from "../components.ts";
import type { AiClient, JsonObject, DocumentModel } from "../domain/types.ts";
import {
  collectDocumentDeriverNames,
  createDefaultDeriverRegistry,
  createDeriverBundle,
  type DocumentDeriverBundle,
  listDeriverDefinitionNames,
} from "../runtime/derivers.ts";
import type { DocumentProject } from "./define.ts";

export interface DocumentProjectManifest {
  schemaVersion: "docxcelerate.project-manifest/v0";
  id: string;
  name: string;
  version: string;
  title: string;
  entrypoint: string;
  builtAt: string;
  previewDocument: string;
  engineDocument: string;
  /**
   * The names these fields had before the vocabulary settled on documents.
   * Still written, because a generation engine reading an artifact is not
   * upgraded at the same moment the toolkit that wrote it is. Read
   * `previewDocument` and `engineDocument`; these go away in a later version.
   *
   * @deprecated
   */
  previewLetter?: string;
  /** @deprecated See {@link DocumentProjectManifest.previewLetter}. */
  engineLetter?: string;
  derivers?: string;
  deriverNames?: string[];
  style?: DocumentModel["style"];
  metadata?: JsonObject;
}

export interface CreateDocumentProjectArtifactOptions {
  entrypoint?: string;
  builtAt?: string;
  previewFileName?: string;
  engineFileName?: string;
  deriversFileName?: string;
}

export interface DocumentProjectArtifact {
  manifest: DocumentProjectManifest;
  previewDocument: DocumentModel;
  engineDocument: DocumentModel;
  /** @deprecated See {@link DocumentProjectManifest.previewLetter}. */
  previewLetter?: DocumentModel;
  /** @deprecated See {@link DocumentProjectManifest.previewLetter}. */
  engineLetter?: DocumentModel;
  derivers?: DocumentDeriverBundle;
}

export interface BuildProjectFinalDocumentOptions<TData = unknown>
  extends Omit<ComponentRuntimeOptions, "dynamicMode"> {
  data?: TData;
}

export async function createDocumentProjectArtifact<TData>(
  project: DocumentProject<TData>,
  options: CreateDocumentProjectArtifactOptions = {},
): Promise<DocumentProjectArtifact> {
  const previewFileName = options.previewFileName ?? "preview.json";
  const engineFileName = options.engineFileName ?? "document.json";
  const deriversFileName = options.deriversFileName ?? "derivers.js";
  const builtAt = options.builtAt ?? new Date().toISOString();
  const previewDocument = await buildProjectPreviewDocument(project);
  const engineDocument = await buildProjectEngineDocument(project);
  const deriverNames = collectDocumentDeriverNames(engineDocument);
  assertReferencedDeriversAreAvailable(project, deriverNames);
  const derivers = createDeriverBundle(project.derivers, {
    names: deriverNames,
    entrypoint: options.entrypoint,
    bundledAt: builtAt,
  });

  return {
    previewDocument,
    engineDocument,
    // Written under both names so an engine that has not been updated still
    // finds what it is looking for.
    previewLetter: previewDocument,
    engineLetter: engineDocument,
    derivers,
    manifest: {
      schemaVersion: "docxcelerate.project-manifest/v0",
      id: project.id,
      name: project.name,
      version: project.version,
      title: project.template.title,
      entrypoint: options.entrypoint ?? "",
      builtAt,
      previewDocument: previewFileName,
      engineDocument: engineFileName,
      previewLetter: previewFileName,
      engineLetter: engineFileName,
      derivers: derivers ? deriversFileName : undefined,
      deriverNames,
      style: project.style,
      metadata: project.metadata,
    },
  };
}

function assertReferencedDeriversAreAvailable<TData>(
  project: DocumentProject<TData>,
  names: string[],
): void {
  const defaultDerivers = createDefaultDeriverRegistry();
  const projectDerivers = new Set(listDeriverDefinitionNames(project.derivers));
  const missing = names.filter((name) => !projectDerivers.has(name) && !defaultDerivers.has(name));

  if (missing.length > 0) {
    throw new Error(
      `Document project "${project.id}" references derivers that are not registered: ${
        missing.join(", ")
      }`,
    );
  }
}

export async function buildProjectPreviewDocument<TData>(
  project: DocumentProject<TData>,
  options: ComponentRuntimeOptions = {},
): Promise<DocumentModel> {
  const doc = await buildDocument(project.template, project.previewData, {
    ...project.previewOptions,
    ...options,
    derivers: project.derivers,
    deriverMode: "resolve",
    dynamicMode: "placeholder",
  });

  return {
    ...doc,
    style: project.style,
    metadata: project.metadata ? { ...project.metadata, ...doc.metadata } : doc.metadata,
  };
}

export async function buildProjectEngineDocument<TData>(
  project: DocumentProject<TData>,
  options: ComponentRuntimeOptions = {},
): Promise<DocumentModel> {
  const data = createEngineArtifactData(project.previewData) as TData;
  const doc = await buildDocument(project.template, data, {
    ...project.buildOptions,
    ...options,
    availableTokens: "{{ctx.availableTokens}}" as unknown as number,
    aiClient: engineArtifactAiClient,
    derivers: project.derivers,
    deriverMode: "preserve",
    dynamicMode: "resolve",
  });

  return {
    ...doc,
    nodes: removeEmptyDynamicText(doc.nodes),
    style: project.style,
    metadata: project.metadata ? { ...project.metadata, ...doc.metadata } : doc.metadata,
  };
}

/**
 * Resolves a project into its final document, running dynamic nodes through the
 * AI client. This is what a generation service calls with request-time data.
 */
export async function buildProjectFinalDocument<TData>(
  project: DocumentProject<TData>,
  options: BuildProjectFinalDocumentOptions<TData> = {},
): Promise<DocumentModel> {
  const { data = project.previewData, ...runtimeOptions } = options;
  const doc = await buildDocument(project.template, data, {
    ...project.buildOptions,
    ...runtimeOptions,
    derivers: project.derivers,
    deriverMode: "resolve",
    dynamicMode: "resolve",
  });

  return {
    ...doc,
    style: project.style,
    metadata: project.metadata ? { ...project.metadata, ...doc.metadata } : doc.metadata,
  };
}

const engineArtifactAiClient: AiClient = {
  generateParagraph(): string {
    return "";
  },
};

function createEngineArtifactData(value: unknown, path: string[] = []): unknown {
  if (Array.isArray(value)) {
    return value.map((item, index) => createEngineArtifactData(item, [...path, String(index)]));
  }

  if (value && typeof value === "object") {
    return Object.fromEntries(
      Object.keys(value).map((key) => [
        key,
        createEngineArtifactData((value as Record<string, unknown>)[key], [...path, key]),
      ]),
    );
  }

  return path.length === 0 ? "{{data.value}}" : `{{data.${path.join(".")}}}`;
}

function removeEmptyDynamicText(nodes: DocumentModel["nodes"]): DocumentModel["nodes"] {
  return nodes.map((node) => {
    if (node.kind === "section") {
      return {
        ...node,
        children: removeEmptyDynamicText(node.children),
      };
    }

    if (node.kind === "paragraph" && node.mode === "dynamic" && node.text === "") {
      const { text: _text, ...withoutText } = node;
      return withoutText;
    }

    return node;
  });
}
