/**
 * Building a document and settling it: the model, the build, the artifact, and
 * the runtime that resolves what a build left open.
 *
 * This is the default entrypoint, and it is the one for the code that *runs*
 * documents — a build step, a generation service, an engine. It is not the
 * widest one and the others are not subsets of it; the entrypoints are split by
 * job, not by size:
 *
 * - `/template` writes documents — the elements, the hooks, `template`.
 * - `/document` is what a document project imports: `/template` plus the model
 *   and `defineDocumentProject`. It is larger than this one.
 * - `/docx` packs a settled model into a Word file.
 * - `/scaffold` and `/cli` create projects; `/registry` and `/registry/install`
 *   describe and install components; `/themes` and `/transform` are themselves.
 *
 * Where two entrypoints carry the same thing, they carry the same thing —
 * `tests/surface.test.ts` fails if they drift.
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
// Reading what an image node points at, which every renderer has to do and no
// renderer should do differently from another.
export * from "./render/image_source.ts";
export * from "./runtime/ai.ts";
// An engine resolving a published document evaluates its conditions, so the
// two functions that do it are part of the surface rather than an internal.
export * from "./runtime/conditions.ts";
export * from "./runtime/data.ts";
export * from "./runtime/deriver_bundle.ts";
export * from "./runtime/deriver_module.ts";
export * from "./runtime/deriver_refs.ts";
export * from "./runtime/derivers.ts";
export * from "./runtime/resolver.ts";
// The catalogs, but not the installer: describing what is in the registry is
// data, and installing it touches a filesystem. Anything that wants the second
// asks for `/registry/install` and says so.
export * from "./registry/mod.ts";
export * from "./theme/mod.ts";
export * from "./version.ts";
