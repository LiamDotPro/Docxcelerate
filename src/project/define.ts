import type { ComponentRuntimeOptions, DocumentTemplate } from "../components.ts";
import type { JsonObject, DocumentStyle } from "../domain/types.ts";
import type { DeriverDefinitions } from "../runtime/derivers.ts";

/**
 * A document project: a template, the data it is previewed against, and
 * everything else a build needs to know.
 *
 * @module
 */

/**
 * A document project, as the build tools and the preview app read it.
 *
 * @typeParam TData The shape the project's template reads.
 */
export interface DocumentProject<TData = unknown> {
  /** The project version, so a reader knows what it is looking at. */
  schemaVersion: "docxcelerate.project/v0";
  /** The project's identifier. */
  id: string;
  /** The project's package name. */
  name: string;
  /** The project's version. */
  version: string;
  /** The template the document is built from. */
  template: DocumentTemplate<TData>;
  /** Stand-in data for the preview, shaped like a real request. */
  previewData: TData;
  /** Computations the document defers to the engine. */
  derivers?: DeriverDefinitions;
  /** How the document looks. */
  style?: DocumentStyle;
  /** Runtime overrides applied when building the preview. */
  previewOptions?: ComponentRuntimeOptions;
  /** Runtime overrides applied when building for an engine. */
  buildOptions?: ComponentRuntimeOptions;
  /** Anything to carry alongside the document. */
  metadata?: JsonObject;
}

/**
 * What {@linkcode defineDocumentProject} takes.
 *
 * The same as {@linkcode DocumentProject} minus the parts it fills in.
 *
 * @typeParam TData The shape the project's template reads.
 */
export interface DefineDocumentProjectOptions<TData> {
  /** The project's identifier. */
  id: string;
  /** The project's package name. */
  name: string;
  /** The project's version. Defaults to `0.1.0`. */
  version?: string;
  /** The template the document is built from. */
  template: DocumentTemplate<TData>;
  /** Stand-in data for the preview, shaped like a real request. */
  previewData: TData;
  /** Computations the document defers to the engine. */
  derivers?: DeriverDefinitions;
  /** How the document looks. */
  style?: DocumentStyle;
  /** Runtime overrides applied when building the preview. */
  previewOptions?: ComponentRuntimeOptions;
  /** Runtime overrides applied when building for an engine. */
  buildOptions?: ComponentRuntimeOptions;
  /** Anything to carry alongside the document. */
  metadata?: JsonObject;
}

/**
 * Declares a document project. This is what a `document.project.ts` exports,
 * and what the CLI and preview app look for.
 *
 * @typeParam TData The shape the project's template reads.
 * @param options The template, its preview data, and how to build it.
 * @returns The project.
 *
 * @example
 * ```ts
 * export default defineDocumentProject<TenancyData>({
 *   id: "tenancy-renewal",
 *   name: "tenancy-renewal",
 *   template: documentTemplate,
 *   previewData,
 * });
 * ```
 */
export function defineDocumentProject<TData>(
  options: DefineDocumentProjectOptions<TData>,
): DocumentProject<TData> {
  return {
    schemaVersion: "docxcelerate.project/v0",
    id: options.id,
    name: options.name,
    version: options.version ?? "0.1.0",
    template: options.template,
    previewData: options.previewData,
    derivers: options.derivers,
    style: options.style,
    previewOptions: options.previewOptions,
    buildOptions: options.buildOptions,
    metadata: options.metadata,
  };
}
