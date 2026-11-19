/**
 * Which demo documents the homepage shows, and which it opens on.
 *
 * A document is listed only once its project exists under src/demo/documents,
 * so one that is still being written stays out of the tree and the hero
 * rather than pointing at a preview that was never built. The order is the
 * order the tree shows them in; the first present one opens.
 */
const projects = import.meta.glob("./demo/documents/*/document.project.ts");

const ORDER = [
  // The portfolio review leads once it lands: thirteen pages, charts, tables,
  // and both kinds of AI part, so it shows the most of the framework.
  "portfolio-review",
  "invoice",
  "repairs-appointment",
  "policy-renewal",
] as const;

export type DemoDocumentId = (typeof ORDER)[number];

export const DEMO_DOCUMENT_IDS: DemoDocumentId[] = ORDER.filter(
  (id) => `./demo/documents/${id}/document.project.ts` in projects,
);

export const OPEN_DOCUMENT_ID: DemoDocumentId = DEMO_DOCUMENT_IDS[0];
