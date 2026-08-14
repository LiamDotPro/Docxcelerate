import { test } from "node:test";
import { assertEquals, assertStringIncludes } from "./assert.ts";
import { cleanMinimalDocumentStyle, type DocumentModel } from "docxcelerate";
import { renderDocumentWebsite } from "docxcelerate/renderer";
import { createDocxBlob } from "docxcelerate/docx";

/**
 * What the renderers do with every node kind, including the ones a build
 * normally resolves away.
 *
 * A renderer that meets an unfamiliar kind and falls through to its default
 * prints the wrong thing rather than failing, so each kind is asserted by name.
 */
const everyKind: DocumentModel = {
  schemaVersion: "docxcelerate.letter/v0",
  id: "kinds",
  title: "Every kind",
  style: cleanMinimalDocumentStyle,
  nodes: [
    { id: "contents", kind: "tableOfContents", title: "What is here" },
    {
      id: "opening",
      kind: "section",
      title: "Opening",
      children: [
        { id: "greeting", kind: "paragraph", mode: "static", text: "Hello Avery," },
        { id: "note", kind: "paragraph", mode: "dynamic", text: "A generated note." },
        { id: "signature", kind: "image", mode: "static", path: "sig.png", alt: "A signature" },
        {
          id: "trend",
          kind: "graph",
          mode: "static",
          graphType: "line",
          data: { values: [1, 2] },
          caption: "A trend",
        },
      ],
    },
    {
      // Only a published document carries one of these, but the preview app can
      // be pointed at a published document, so both renderers meet it.
      id: "visits",
      kind: "repeat",
      source: { scope: "data", path: "visits" },
      as: "visit",
      indexAs: "index",
      children: [
        { id: "visit", kind: "paragraph", mode: "static", text: "A visit." },
      ],
    },
  ],
};

test("the web renderer prints every node kind rather than falling through", () => {
  const html = renderDocumentWebsite(everyKind);

  assertStringIncludes(html, "Hello Avery,");
  assertStringIncludes(html, "A generated note.");
  assertStringIncludes(html, "A signature");
  assertStringIncludes(html, "A trend");
  assertStringIncludes(html, "What is here");
  // The loop says it is a loop, and shows its body.
  assertStringIncludes(html, "Repeats per visits");
  assertStringIncludes(html, "A visit.");
});

test("the web renderer marks a dynamic paragraph as dynamic", () => {
  const html = renderDocumentWebsite(everyKind);

  assertStringIncludes(html, "paragraph-dynamic");
  assertStringIncludes(html, "paragraph-static");
});

test("the web renderer escapes text rather than letting it become markup", () => {
  const html = renderDocumentWebsite({
    ...everyKind,
    nodes: [{
      id: "hostile",
      kind: "paragraph",
      mode: "static",
      text: '<script>alert("x")</script>',
    }],
  });

  assertEquals(html.includes("<script>alert"), false);
  assertStringIncludes(html, "&lt;script&gt;");
});

test("the DOCX packer accepts every node kind and produces a file", async () => {
  const blob = await createDocxBlob(everyKind);

  assertEquals(blob.size > 0, true);
  // A .docx is a zip, so the first two bytes are the local file header.
  const header = new Uint8Array(await blob.slice(0, 2).arrayBuffer());
  assertEquals([header[0], header[1]], [0x50, 0x4b]);
});

test("a repeat body reaches the DOCX rather than being dropped", async () => {
  const withLoop = await createDocxBlob(everyKind);
  const withoutLoop = await createDocxBlob({
    ...everyKind,
    nodes: everyKind.nodes.filter((node) => node.kind !== "repeat"),
  });

  // The body is one more paragraph of content, so the packed file differs.
  assertEquals(withLoop.size === withoutLoop.size, false);
});
