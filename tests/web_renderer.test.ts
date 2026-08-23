import { test } from "node:test";
import { assertStringIncludes } from "./assert.ts";
import type { DocumentModel } from "docxcelerate";
import { renderDocumentWebsite } from "docxcelerate/renderer";

test("web renderer creates a Word-style workspace with one A4 page", () => {
  const letter: DocumentModel = {
    schemaVersion: "docxcelerate.letter/v0",
    id: "sample",
    title: "Sample Letter",
    nodes: [
      {
        id: "opening",
        kind: "section",
        title: "Opening",
        children: [
          {
            id: "greeting",
            kind: "paragraph",
            mode: "static",
            text: "Hello Avery,",
          },
        ],
      },
    ],
  };

  const html = renderDocumentWebsite(letter);

  assertStringIncludes(html, 'class="workspace"');
  assertStringIncludes(html, 'class="workspace-inner"');
  assertStringIncludes(html, 'class="a4-page"');
  // The page is sized from the document's style rather than from constants in
  // the renderer, so A4 arrives as custom properties. A document with no style
  // of its own — this one — falls back to the default theme, which is A4.
  assertStringIncludes(html, "--page-width: 210mm;");
  assertStringIncludes(html, "--page-height: 297mm;");
  assertStringIncludes(html, "min-height: var(--page-height)");
  assertStringIncludes(html, "min-width: 1180px");
});
