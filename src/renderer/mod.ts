/**
 * The renderers: a finished document out to DOCX bytes, or to a page you can
 * read in a browser.
 *
 * Both take a settled `DocumentModel` — whatever produced it, and
 * whether it was built locally or by an engine, is not their concern.
 *
 * @module
 */

export * from "./docx_renderer.ts";
export * from "./web_renderer.ts";
