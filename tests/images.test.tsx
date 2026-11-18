import { test } from "node:test";
import { assertEquals, assertStringIncludes } from "./assert.ts";
import { buildDocument, imageSourceOf, isSvg, pixelSizeOf, rasterTypeOf } from "docxcelerate";
import { createDocxBlob, createDocxDocument } from "docxcelerate/docx";
import { documentXml } from "./docx.ts";

import { Document, Image, Paragraph, template } from "docxcelerate/template";

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
// A side left out follows the picture's own proportions
// ---------------------------------------------------------------------------

// 4x2 and 2x4 PNGs: one wide, one tall, so a swapped width/height read fails.
const pngWide = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAIAAADwyuo0AAAAC0lEQVR42mNgwAUAABoAAS+Yl6YAAAAASUVORK5CYII=";
const pngTall = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAAECAIAAAArjXluAAAAC0lEQVR42mNgwA0AABwAAcdGaXkAAAAASUVORK5CYII=";
const gifWide = "data:image/gif;base64,R0lGODdhBAACAIAAAAAAAP///yH5BAQAAAAALAAAAAAEAAIAAAIDhG8FADs=";

function bytesOf(uri: string): Uint8Array {
  const source = imageSourceOf(uri);
  if (source.kind !== "data") throw new Error("not a data URI");
  return source.bytes;
}

const jpeg = (sof: number) =>
  new Uint8Array([
    0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x04, 0x00, 0x00, 0xFF, 0xC4, 0x00, 0x04, 0x00, 0x00,
    0xFF, sof, 0x00, 0x11, 0x08, 0x00, 0x02, 0x00, 0x04, 0x03,
  ]);

test("PNG and GIF headers give their pixel size", () => {
  assertEquals(pixelSizeOf(bytesOf(pngWide)), { width: 4, height: 2 });
  assertEquals(pixelSizeOf(bytesOf(pngTall)), { width: 2, height: 4 });
  assertEquals(pixelSizeOf(bytesOf(gifWide)), { width: 4, height: 2 });
});

test("a JPEG's size is found past segments that are not the frame header", () => {
  // The DHT (C4) before the SOF must not be mistaken for it.
  assertEquals(pixelSizeOf(jpeg(0xC0)), { width: 4, height: 2 });
  assertEquals(pixelSizeOf(jpeg(0xC2)), { width: 4, height: 2 });
});

test("a JPEG with no frame header has no size", () => {
  assertEquals(pixelSizeOf(jpeg(0xC0).slice(0, 8)), undefined);
  assertEquals(pixelSizeOf(new Uint8Array([0xFF, 0xD8, 0xFF, 0xDA, 0x00, 0x02])), undefined);
});

test("a top-down BMP's negative height is still a height", () => {
  const bmp = new Uint8Array(26);
  bmp.set([0x42, 0x4D]);
  new DataView(bmp.buffer).setInt32(18, 4, true);
  new DataView(bmp.buffer).setInt32(22, -2, true);
  assertEquals(pixelSizeOf(bmp), { width: 4, height: 2 });
});

test("bytes that are not a sized image have no size", () => {
  assertEquals(pixelSizeOf(new Uint8Array([0, 0, 0])), undefined);
  assertEquals(pixelSizeOf(new Uint8Array()), undefined);
  const zero = bytesOf(pngWide).slice();
  zero[19] = 0;
  assertEquals(pixelSizeOf(zero), undefined);
});

async function extentOf(node: unknown): Promise<string> {
  return await documentXml(await build(<Document id="d" title="D">{node as never}</Document>));
}

test("width alone keeps the proportions, and so does height alone", async () => {
  const extent = '<wp:extent cx="2028825" cy="1019175"/>';
  assertStringIncludes(await extentOf(<Image id="m" src={pngWide} alt="M" width={160} />), extent);
  assertStringIncludes(await extentOf(<Image id="m" src={pngWide} alt="M" height={80} />), extent);
});

test("a tall picture follows its own shape", async () => {
  assertStringIncludes(
    await extentOf(<Image id="m" src={pngTall} alt="M" width={60} />),
    '<wp:extent cx="762000" cy="1524000"/>',
  );
});

test("neither side is 120pt wide at the picture's height", async () => {
  assertStringIncludes(
    await extentOf(<Image id="m" src={pngWide} alt="M" />),
    '<wp:extent cx="1524000" cy="762000"/>',
  );
});

test("both sides given are drawn as given", async () => {
  assertStringIncludes(
    await extentOf(<Image id="m" src={pngWide} alt="M" width={160} height={160} />),
    '<wp:extent cx="2028825" cy="2028825"/>',
  );
});

test("an SVG follows the proportions of its fallback raster", async () => {
  assertStringIncludes(
    await extentOf(<Image id="m" src={svgMark} fallbackSrc={pngWide} alt="M" width={160} />),
    '<wp:extent cx="2028825" cy="1019175"/>',
  );
});

test("a header that says nothing stays square", async () => {
  assertStringIncludes(
    await extentOf(<Image id="m" src="data:image/png;base64,AAAA" alt="M" width={100} />),
    '<wp:extent cx="1266825" cy="1266825"/>',
  );
});

test("a picture inside a paragraph is sized the same way", async () => {
  assertStringIncludes(
    await extentOf(
      <Paragraph id="p">
        <Image id="m" src={pngWide} alt="M" width={160} />
        Logo
      </Paragraph>,
    ),
    '<wp:extent cx="2028825" cy="1019175"/>',
  );
});

test("the worked-out size is never written into the model", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Image id="m" src={pngWide} width={160} />
    </Document>,
  );
  const node = doc.nodes[0];

  assertEquals(node.kind === "image" && node.width, 160);
  assertEquals(node.kind === "image" && node.height, undefined);
});
