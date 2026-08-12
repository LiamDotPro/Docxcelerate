import type { ComponentRuntimeOptions, LetterTemplate } from "../components.ts";
import type { JsonObject, LetterStyle } from "../domain/types.ts";
import type { DeriverDefinitions } from "../runtime/derivers.ts";

export interface LetterProject<TData = unknown> {
  schemaVersion: "docxcelerate.project/v0";
  id: string;
  name: string;
  version: string;
  template: LetterTemplate<TData>;
  previewData: TData;
  derivers?: DeriverDefinitions;
  style?: LetterStyle;
  previewOptions?: ComponentRuntimeOptions;
  buildOptions?: ComponentRuntimeOptions;
  metadata?: JsonObject;
}

export interface DefineLetterProjectOptions<TData> {
  id: string;
  name: string;
  version?: string;
  template: LetterTemplate<TData>;
  previewData: TData;
  derivers?: DeriverDefinitions;
  style?: LetterStyle;
  previewOptions?: ComponentRuntimeOptions;
  buildOptions?: ComponentRuntimeOptions;
  metadata?: JsonObject;
}

export function defineLetterProject<TData>(
  options: DefineLetterProjectOptions<TData>,
): LetterProject<TData> {
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
