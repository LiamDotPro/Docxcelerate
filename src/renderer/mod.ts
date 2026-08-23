/**
 * The renderer: a finished document out to the `.docx` bytes it is for.
 *
 * It takes a settled `DocumentModel` — whatever produced it, and whether it
 * was built locally or by an engine, is not its concern.
 *
 * To show a document on a screen, pack it and read the file back with a DOCX
 * viewer such as `docx-preview`, which is what a scaffolded workspace's preview
 * app does. A second renderer that laid the model out in HTML would be a second
 * answer to what the document looks like, and the one nobody opens in Word is
 * the one that drifts.
 *
 * @module
 */

export * from "./docx_renderer.ts";
