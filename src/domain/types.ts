/**
 * The published document model: the shape a built document takes once the JSX
 * has been evaluated and before an engine turns it into a DOCX.
 *
 * Everything here is data. A `DocumentModel` is JSON — it carries no closures
 * and no component code — which is what lets a document be built on one machine
 * and rendered on another, at a different time, against data that did not exist
 * at build time. The decisions that cannot be made until then survive as
 * {@linkcode Condition}, {@linkcode DeriverInvocation} and {@linkcode RepeatNode}.
 *
 * The model is one import because that is how it is read; it is four files
 * because that is how it is written.
 *
 * @module
 */

export * from "./expressions.ts";
export * from "./nodes.ts";
export * from "./runtime.ts";
export * from "./style.ts";
