import { buildDocument, type ComponentRuntimeOptions } from "../components.ts";
import type { AiClient, JsonObject, DocumentModel } from "../domain/types.ts";
import {
  collectDocumentDeriverNames,
  createDeriverBundle,
  type DocumentDeriverBundle,
} from "../runtime/deriver_bundle.ts";
import {
  createDefaultDeriverRegistry,
  listDeriverDefinitionNames,
} from "../runtime/derivers.ts";
import { createPublishData } from "../template/publish.ts";
import type { DocumentProject } from "./define.ts";

/**
 * Turning a document project into the files that leave the machine it was
 * written on.
 *
 * A build produces two documents from one template. The preview is fully
 * resolved against sample data, for reading on a screen. The engine document
 * keeps every decision open — branches, loops and tokens intact — because the
 * data that settles them does not exist yet.
 *
 * @module
 */

/** What an artifact is, described for whatever reads it later. */
export interface DocumentProjectManifest {
  /** The manifest version, so a reader knows what it is looking at. */
  schemaVersion: "docxcelerate.project-manifest/v0";
  /** The project's identifier. */
  id: string;
  /** The project's package name. */
  name: string;
  /** The project's version. */
  version: string;
  /** The document's title. */
  title: string;
  /** The project file the artifact was built from. */
  entrypoint: string;
  /** When the artifact was built, as an ISO 8601 string. */
  builtAt: string;
  /** The file holding the preview document. */
  previewDocument: string;
  /** The file holding the engine document. */
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
  /**
   * The file holding the engine document, under its former name.
   *
   * @deprecated See {@link DocumentProjectManifest.previewLetter}.
   */
  engineLetter?: string;
  /** The file holding the deriver bundle, when the document uses any. */
  derivers?: string;
  /** The derivers the engine document invokes. */
  deriverNames?: string[];
  /** How the document looks. */
  style?: DocumentModel["style"];
  /** Anything the project carries alongside the document. */
  metadata?: JsonObject;
}

/** What {@linkcode createDocumentProjectArtifact} takes beyond the project. */
export interface CreateDocumentProjectArtifactOptions {
  /** The project file being built, recorded on the manifest. */
  entrypoint?: string;
  /** The build timestamp to record. Defaults to now. */
  builtAt?: string;
  /** What to call the preview document. Defaults to `preview.json`. */
  previewFileName?: string;
  /** What to call the engine document. Defaults to `document.json`. */
  engineFileName?: string;
  /** What to call the deriver bundle. Defaults to `derivers.js`. */
  deriversFileName?: string;
}

/** Everything a build produces, ready to be written out or shipped. */
export interface DocumentProjectArtifact {
  /** What the artifact is. */
  manifest: DocumentProjectManifest;
  /** The document resolved against the project's preview data. */
  previewDocument: DocumentModel;
  /** The document with its decisions still open, for an engine. */
  engineDocument: DocumentModel;
  /**
   * The preview document, under its former name.
   *
   * @deprecated See {@link DocumentProjectManifest.previewLetter}.
   */
  previewLetter?: DocumentModel;
  /**
   * The engine document, under its former name.
   *
   * @deprecated See {@link DocumentProjectManifest.previewLetter}.
   */
  engineLetter?: DocumentModel;
  /** The project's derivers as loadable source, when the document uses any. */
  derivers?: DocumentDeriverBundle;
}

/**
 * What {@linkcode buildProjectFinalDocument} takes beyond the project.
 *
 * @typeParam TData The shape the project's template reads.
 */
export interface BuildProjectFinalDocumentOptions<TData = unknown>
  extends Omit<ComponentRuntimeOptions, "dynamicMode"> {
  /** The data to write this document from. Falls back to the preview data. */
  data?: TData;
}

/**
 * Builds a project into everything that ships: both documents, the manifest,
 * and the deriver bundle.
 *
 * @typeParam TData The shape the project's template reads.
 * @param project The project to build.
 * @param options File names and build metadata to record.
 * @returns The artifact.
 * @throws If the document invokes a deriver the project does not register.
 */
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

/**
 * Builds the document you read on a screen.
 *
 * A preview is rebuilt every time a file is saved, so it is built to be quick
 * before it is built to be complete. Anything it waits for is time a person
 * spends watching a document fail to appear, and a preview nobody waits for is
 * one they keep open while they write.
 *
 * So the two things that cost time stand in rather than run. A generated node
 * shows the placeholder `useAi` required, instead of calling a model. A deriver
 * that declared a stand-in — a code to render, a file to read, a service to ask
 * — uses it, while the cheap ones still run so the figures on the page are the
 * real ones. Both are resolved for real when a document is actually written,
 * which is the moment waiting is worth something.
 *
 * @typeParam TData The shape the project's template reads.
 * @param project The project to build.
 * @param options Runtime overrides for this build.
 * @returns The resolved document.
 */
export async function buildProjectPreviewDocument<TData>(
  project: DocumentProject<TData>,
  options: ComponentRuntimeOptions = {},
): Promise<DocumentModel> {
  const doc = await buildDocument(project.template, project.previewData, {
    ...project.previewOptions,
    ...options,
    derivers: project.derivers,
    deriverMode: "placeholder",
    dynamicMode: "placeholder",
  });

  return {
    ...doc,
    style: project.style,
    metadata: project.metadata ? { ...project.metadata, ...doc.metadata } : doc.metadata,
  };
}

/**
 * Builds the document that goes to an engine.
 *
 * Nothing here is decided. The data is a stand-in for a request nobody has made,
 * so a branch keeps both arms under the condition that selects one, a loop stays
 * a loop, and every value stays the token the engine substitutes. What comes out
 * is the whole document with the decisions still in it.
 *
 * @typeParam TData The shape the project's template reads.
 * @param project The project to build.
 * @param options Runtime overrides for this build.
 * @returns The document, with its decisions still open.
 */
export async function buildProjectEngineDocument<TData>(
  project: DocumentProject<TData>,
  options: ComponentRuntimeOptions = {},
): Promise<DocumentModel> {
  const doc = await buildDocument(project.template, createPublishData() as TData, {
    ...project.buildOptions,
    ...options,
    availableTokens: "{{ctx.availableTokens}}" as unknown as number,
    aiClient: engineArtifactAiClient,
    derivers: project.derivers,
    deriverMode: "preserve",
    dynamicMode: "resolve",
    branchMode: "publish",
  });

  return {
    ...doc,
    nodes: removeEmptyDynamicText(doc.nodes),
    header: doc.header && removeEmptyDynamicText(doc.header),
    footer: doc.footer && removeEmptyDynamicText(doc.footer),
    firstHeader: doc.firstHeader && removeEmptyDynamicText(doc.firstHeader),
    firstFooter: doc.firstFooter && removeEmptyDynamicText(doc.firstFooter),
    style: project.style,
    metadata: project.metadata ? { ...project.metadata, ...doc.metadata } : doc.metadata,
  };
}

/**
 * Resolves a project into its final document, running dynamic nodes through the
 * AI client. This is what a generation service calls with request-time data.
 *
 * @typeParam TData The shape the project's template reads.
 * @param project The project to write a document from.
 * @param options The request's data, and runtime overrides for this build.
 * @returns The finished document, ready to render.
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

function removeEmptyDynamicText(nodes: DocumentModel["nodes"]): DocumentModel["nodes"] {
  return nodes.map((node) => {
    if (node.kind === "section" || node.kind === "repeat") {
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
