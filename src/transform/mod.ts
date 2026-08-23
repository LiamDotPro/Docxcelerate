/**
 * The compiler that lets a document be written in ordinary TypeScript.
 *
 * This is a build-time entrypoint. Nothing here runs while a document is being
 * written, so `typescript` is a peer dependency rather than a runtime one — a
 * service that only renders published documents never loads any of it.
 *
 * @module
 */

export {
  assertCompiledSources,
  compiledMarker,
  findUncompiledSources,
  isCompiledSource,
  type TransformOptions,
  type TransformResult,
  transformDocumentSource,
  type UncompiledSource,
} from "./compile.ts";
export {
  docxcelerateEsbuildTransform,
  docxcelerateTransform,
  type DocxcelerateTransformPluginOptions,
  type EsbuildBuildLike,
  type EsbuildPluginLike,
  type VitePluginLike,
} from "./plugins.ts";
