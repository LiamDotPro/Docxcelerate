import type { ComponentRuntimeOptions, DocumentTemplate } from "../components.ts";
import type { JsonObject, DocumentStyle } from "../domain/types.ts";
import type { DeriverDefinitions } from "../runtime/derivers.ts";

export interface DocumentProject<TData = unknown> {
  schemaVersion: "docxcelerate.project/v0";
  id: string;
  name: string;
  version: string;
  template: DocumentTemplate<TData>;
  previewData: TData;
  derivers?: DeriverDefinitions;
  style?: DocumentStyle;
  previewOptions?: ComponentRuntimeOptions;
  buildOptions?: ComponentRuntimeOptions;
  metadata?: JsonObject;
}

export interface DefineDocumentProjectOptions<TData> {
  id: string;
  name: string;
  version?: string;
  template: DocumentTemplate<TData>;
  previewData: TData;
  derivers?: DeriverDefinitions;
  style?: DocumentStyle;
  previewOptions?: ComponentRuntimeOptions;
  buildOptions?: ComponentRuntimeOptions;
  metadata?: JsonObject;
}

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
