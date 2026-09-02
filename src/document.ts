/**
 * What a document project needs and nothing else: the model, the components,
 * the template surface, and `defineDocumentProject`.
 *
 * This is `/template` plus everything a `document.project.ts` names — so a
 * node imports `/template` and the project file beside it imports this, and
 * neither has to reach for the build, the packer or the scaffolder to say
 * what the document is.
 *
 * @module
 */

export * from "./domain/types.ts";
export * from "./components.ts";
export * from "./project/define.ts";
export * from "./project/style.ts";
export * from "./runtime/deriver_refs.ts";
export * from "./runtime/derivers.ts";
export * from "./template.ts";
