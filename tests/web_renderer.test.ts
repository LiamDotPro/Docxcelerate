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
  assertStringIncludes(html, "width: 210mm");
  assertStringIncludes(html, "min-height: 297mm");
  assertStringIncludes(html, "min-width: 1180px");
});
