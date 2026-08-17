/**
 * The whole toolkit: the document model, the components that build one, and the
 * runtime that settles it against data.
 *
 * This is the default entrypoint, and the widest one. The narrower entrypoints
 * are subsets of it — `/template` for writing documents, `/docx` for packing
 * them, `/scaffold` for creating projects.
 *
 * @example Building a project and packing the result
 * ```ts
 * import { buildProjectFinalDocument } from "@docxcelerate/docxcelerate";
 * import { createDocxBlob } from "@docxcelerate/docxcelerate/docx";
 * import project from "./documents/welcome/document.project.ts";
 *
 * const doc = await buildProjectFinalDocument(project, { data: { name: "Avery" } });
 * const blob = await createDocxBlob(doc);
 * ```
 *
 * @module
 */

export * from "./domain/types.ts";
export * from "./components.ts";
export * from "./project/artifact.ts";
export * from "./project/define.ts";
export * from "./project/style.ts";
export * from "./runtime/ai.ts";
// An engine resolving a published document evaluates its conditions, so the
// two functions that do it are part of the surface rather than an internal.
export * from "./runtime/conditions.ts";
export * from "./runtime/data.ts";
export * from "./runtime/derivers.ts";
export * from "./runtime/resolver.ts";
