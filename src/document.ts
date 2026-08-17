/**
 * What a document project needs and nothing else: the model, the components,
 * the template surface, and `defineDocumentProject`.
 *
 * A narrower `/` — it leaves out the scaffolding and the renderers, so a
 * document imports only what it actually composes with.
 *
 * @module
 */

export * from "./domain/types.ts";
export * from "./components.ts";
export * from "./project/define.ts";
export * from "./project/style.ts";
export * from "./runtime/derivers.ts";
export * from "./template.ts";
