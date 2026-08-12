import { buildLetterDocument, type ComponentRuntimeOptions } from "../components.ts";
import type { AiClient, JsonObject, LetterDocument } from "../domain/types.ts";
import {
  collectLetterDeriverNames,
  createDefaultDeriverRegistry,
  createDeriverBundle,
  type LetterDeriverBundle,
  listDeriverDefinitionNames,
} from "../runtime/derivers.ts";
import type { LetterProject } from "./define.ts";

export interface LetterProjectManifest {
  schemaVersion: "docxcelerate.project-manifest/v0";
  id: string;
  name: string;
  version: string;
  title: string;
  entrypoint: string;
  builtAt: string;
  previewLetter: string;
  engineLetter: string;
  derivers?: string;
  deriverNames?: string[];
  style?: LetterDocument["style"];
  metadata?: JsonObject;
}

export interface CreateLetterProjectArtifactOptions {
  entrypoint?: string;
  builtAt?: string;
  previewFileName?: string;
  engineFileName?: string;
  deriversFileName?: string;
}

export interface LetterProjectArtifact {
  manifest: LetterProjectManifest;
  previewLetter: LetterDocument;
  engineLetter: LetterDocument;
  derivers?: LetterDeriverBundle;
}

export async function createLetterProjectArtifact<TData>(
  project: LetterProject<TData>,
  options: CreateLetterProjectArtifactOptions = {},
): Promise<LetterProjectArtifact> {
  const previewFileName = options.previewFileName ?? "preview.json";
  const engineFileName = options.engineFileName ?? "letter.json";
  const deriversFileName = options.deriversFileName ?? "derivers.js";
  const builtAt = options.builtAt ?? new Date().toISOString();
  const previewLetter = await buildProjectPreviewLetter(project);
  const engineLetter = await buildProjectEngineLetter(project);
  const deriverNames = collectLetterDeriverNames(engineLetter);
  assertReferencedDeriversAreAvailable(project, deriverNames);
  const derivers = createDeriverBundle(project.derivers, {
    names: deriverNames,
    entrypoint: options.entrypoint,
    bundledAt: builtAt,
  });

  return {
    previewLetter,
    engineLetter,
    derivers,
    manifest: {
      schemaVersion: "docxcelerate.project-manifest/v0",
      id: project.id,
      name: project.name,
      version: project.version,
      title: project.template.title,
      entrypoint: options.entrypoint ?? "",
      builtAt,
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
  project: LetterProject<TData>,
  names: string[],
): void {
  const defaultDerivers = createDefaultDeriverRegistry();
  const projectDerivers = new Set(listDeriverDefinitionNames(project.derivers));
  const missing = names.filter((name) => !projectDerivers.has(name) && !defaultDerivers.has(name));

  if (missing.length > 0) {
    throw new Error(
      `Letter project "${project.id}" references derivers that are not registered: ${
        missing.join(", ")
      }`,
    );
  }
}

export async function buildProjectPreviewLetter<TData>(
  project: LetterProject<TData>,
  options: ComponentRuntimeOptions = {},
): Promise<LetterDocument> {
  const letter = await buildLetterDocument(project.template, project.previewData, {
    ...project.previewOptions,
    ...options,
    derivers: project.derivers,
    deriverMode: "resolve",
    dynamicMode: "placeholder",
  });

  return {
    ...letter,
    style: project.style,
    metadata: project.metadata ? { ...project.metadata, ...letter.metadata } : letter.metadata,
  };
}

export async function buildProjectEngineLetter<TData>(
  project: LetterProject<TData>,
  options: ComponentRuntimeOptions = {},
): Promise<LetterDocument> {
  const data = createEngineArtifactData(project.previewData) as TData;
  const letter = await buildLetterDocument(project.template, data, {
    ...project.buildOptions,
    ...options,
    availableTokens: "{{ctx.availableTokens}}" as unknown as number,
    aiClient: engineArtifactAiClient,
    derivers: project.derivers,
    deriverMode: "preserve",
    dynamicMode: "resolve",
  });

  return {
    ...letter,
    nodes: removeEmptyDynamicText(letter.nodes),
    style: project.style,
    metadata: project.metadata ? { ...project.metadata, ...letter.metadata } : letter.metadata,
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

function removeEmptyDynamicText(nodes: LetterDocument["nodes"]): LetterDocument["nodes"] {
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
