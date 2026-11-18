import { test } from "node:test";
import { assertEquals, assertStringIncludes } from "./assert.ts";
import { buildDocument, cleanMinimalDocumentStyle, imageSourceOf, isSvg, rasterTypeOf } from "docxcelerate";
import { documentXml, elementOf } from "./docx.ts";
import { createDocxBlob, createDocxDocument } from "docxcelerate/docx";

import { Cell, Document, Image, Row, Table, template } from "docxcelerate/template";

/**
 * What a document can actually show.
 *
 * The rule the whole module turns on: a model is JSON that travels, so the only
 * picture a renderer can be sure of is one carried in the model. A `data:` URI
 * is that; a path is a promise about a machine the document may never be
 * written on. Getting this wrong is quiet — a build that reached for a file
 * would work here and fail wherever the engine runs.
 */

// A one-pixel PNG, small enough to read: the shortest thing that is really an
// image rather than a string that looks like one.
const pngPixel = "data:image/png;base64," +
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

const svgMark = "data:image/svg+xml;base64," +
  "PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSI4IiBoZWlnaHQ9IjgiPjwvc3ZnPg==";

function build(node: unknown) {
  return buildDocument(
    template<Record<string, never>>(node as never),
    {},
    { branchMode: "decide", dynamicMode: "placeholder" },
  );
}

// ---------------------------------------------------------------------------
// Reading what a node points at
// ---------------------------------------------------------------------------

test("a data URI is read into the bytes it carries", () => {
  const source = imageSourceOf(pngPixel);

  assertEquals(source.kind, "data");
  assertEquals(source.kind === "data" && source.mediaType, "image/png");
  // The PNG signature, which is how we know these are really the image's bytes.
  assertEquals(
    source.kind === "data" && Array.from(source.bytes.slice(0, 4)),
    [137, 80, 78, 71],
  );
});

test("a path is a URL, not bytes — nothing here goes looking for it", () => {
  assertEquals(imageSourceOf("./logo.png").kind, "url");
  assertEquals(imageSourceOf("https://example.com/logo.png").kind, "url");
});

test("no path at all is nothing to draw", () => {
  assertEquals(imageSourceOf(undefined).kind, "none");
  assertEquals(imageSourceOf("").kind, "none");
});

test("a data URI that will not decode is nothing to draw, rather than a crash", () => {
  // Somebody pastes half a URI in. The node falls back to its placeholder; the
  // build does not fail on a string.
  assertEquals(imageSourceOf("data:image/png;base64,!!!not base64!!!").kind, "none");
  assertEquals(imageSourceOf("data:image/png").kind, "none");
});

test("the media types Word takes directly are the raster ones", () => {
  assertEquals(rasterTypeOf("image/png"), "png");
  assertEquals(rasterTypeOf("image/jpeg"), "jpg");
  assertEquals(rasterTypeOf("image/svg+xml"), undefined);
  assertEquals(isSvg("image/svg+xml"), true);
});

// ---------------------------------------------------------------------------
// What a node says about a picture it does not have yet
// ---------------------------------------------------------------------------

test("a node with no picture yet says what will stand there, and how big", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Image
        id="qr"
        generalPrompt="A payment QR code."
        placeholder="Scan-to-pay code"
        alt="Scan to pay"
        width={108}
        height={108}
      />
    </Document>,
  );

  const node = doc.nodes[0];

  // The size and the description belong to the node, not to the picture an
  // engine will draw into it. Dropping them left whatever renders the document
  // reserving no room for a picture whose dimensions the template had stated.
  assertEquals(node.kind === "image" ? node.width : undefined, 108);
  assertEquals(node.kind === "image" ? node.height : undefined, 108);
  assertEquals(node.kind === "image" ? node.alt : undefined, "Scan to pay");
  assertEquals(node.kind === "image" ? node.placeholder : undefined, "Scan-to-pay code");
});

// ---------------------------------------------------------------------------
// Into Word
// ---------------------------------------------------------------------------

test("a raster carried in the model is embedded in the file", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Image id="mark" src={pngPixel} alt="The mark" width={8} height={8} />
    </Document>,
  );

  const blob = await createDocxBlob(doc);
  const bytes = new Uint8Array(await blob.arrayBuffer());

  // A .docx is a zip. An embedded image becomes a part in word/media/, so the
  // name appears in the archive's directory — which is how we know the bytes
  // went in rather than a note about them.
  assertStringIncludes(new TextDecoder().decode(bytes), "word/media/");
});

test("an SVG with a raster beside it packs, because Word takes the pair", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Image id="mark" src={svgMark} fallbackSrc={pngPixel} alt="Mark" width={8} height={8} />
    </Document>,
  );

  assertEquals(typeof createDocxDocument(doc), "object");
  assertEquals(doc.nodes[0].kind === "image" && doc.nodes[0].fallbackPath, pngPixel);
});

test("an SVG with nothing beside it embeds nothing, rather than an empty frame", async () => {
  // Word will not embed an SVG alone. With no raster to fall back to the node
  // packs as a note saying what the picture was, so no media part is written
  // and the file does not carry a blank box where a logo should be.
  const doc = await build(
    <Document id="d" title="D">
      <Image id="mark" src={svgMark} alt="Mark" />
    </Document>,
  );

  const blob = await createDocxBlob(doc);
  const text = new TextDecoder().decode(new Uint8Array(await blob.arrayBuffer()));

  assertEquals(text.includes("word/media/"), false);
});

test("a path is not fetched while packing, so the same model packs anywhere", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Image id="mark" src="./logo.png" alt="Logo" />
    </Document>,
  );

  const blob = await createDocxBlob(doc);
  const text = new TextDecoder().decode(new Uint8Array(await blob.arrayBuffer()));

  assertEquals(text.includes("word/media/"), false);
});

// ---------------------------------------------------------------------------
// Where the picture sits
// ---------------------------------------------------------------------------

test("align is on the node when it is said, and absent when it is not", async () => {
  const said = await build(<Document id="d" title="D"><Image id="m" src={pngPixel} align="center" /></Document>);
  const unsaid = await build(<Document id="d" title="D"><Image id="m" src={pngPixel} /></Document>);

  assertEquals(said.nodes[0].kind === "image" && said.nodes[0].align, "center");
  assertEquals("align" in unsaid.nodes[0], false);
});

test("a centred or right-ranged picture sits on a line that says so", async () => {
  for (const align of ["center", "right"] as const) {
    const doc = await build(<Document id="d" title="D"><Image id="m" src={pngPixel} align={align} /></Document>);

    assertStringIncludes(elementOf(await documentXml(doc), "w:p", "<w:drawing>"), `<w:jc w:val="${align}"/>`);
  }
});

test("a picture with no align writes no alignment at all", async () => {
  const doc = await build(<Document id="d" title="D"><Image id="m" src={pngPixel} /></Document>);

  assertEquals(elementOf(await documentXml(doc), "w:p", "<w:drawing>").includes("<w:jc"), false);
});

test("a placeholder sits where the picture will", async () => {
  const doc = await build(<Document id="d" title="D"><Image id="m" src="./logo.png" alt="Logo" align="center" /></Document>);

  assertStringIncludes(elementOf(await documentXml(doc), "w:p", "[image: Logo]"), '<w:jc w:val="center"/>');
});

test("a card moves as a whole, because the paragraph inside it cannot be seen", async () => {
  const doc = {
    ...await build(<Document id="d" title="D"><Image id="m" variant="card" align="center" src={pngPixel} /></Document>),
    style: { ...cleanMinimalDocumentStyle, blocks: { card: { border: "D1D5DB", paddingPt: 4 } } },
  };
  const table = elementOf(await documentXml(doc), "w:tbl", "<w:drawing>");

  assertStringIncludes(table.match(/<w:tblPr>[\s\S]*?<\/w:tblPr>/)?.[0] ?? "", '<w:jc w:val="center"/>');
});

test("inside a cell the column governs, not the picture", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Table id="t" columns={[{ width: "auto", align: "center" }]}>
        <Row><Cell><Image id="m" align="right" src={pngPixel} /></Cell></Row>
      </Table>
    </Document>,
  );
  const paragraph = elementOf(await documentXml(doc), "w:p", "<w:drawing>");

  assertStringIncludes(paragraph, '<w:jc w:val="center"/>');
  assertEquals(paragraph.includes('w:val="right"'), false);
});
