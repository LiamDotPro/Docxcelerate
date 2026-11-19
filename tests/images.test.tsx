import { test } from "node:test";
import { assertEquals, assertStringIncludes } from "./assert.ts";
import { buildDocument, cleanMinimalDocumentStyle, imageSourceOf, isSvg, pixelSizeOf, rasterTypeOf } from "docxcelerate";
import { createDocxBlob, createDocxDocument } from "docxcelerate/docx";
import { documentXml, elementOf, entryOf, paragraphOf, partXml } from "./docx.ts";

import { Cell, createPublishData, Document, Image, Paragraph, Row, Shape, Table, template, useState } from "docxcelerate/template";
import { createDeriverRegistry, EchoAiClient, InMemoryDataProvider, resolveDocument } from "docxcelerate";
import type { ImageNode, ParagraphNode, RuntimeState } from "docxcelerate";
import { settleDocxPreview } from "docxcelerate/preview";
import { JSDOM } from "jsdom";

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

const columnPage = (over = {}) => ({
  size: "A4",
  orientation: "portrait",
  margins: { topMm: 20, rightMm: 20, bottomMm: 20, leftMm: 20 },
  ...over,
});

async function styled(node: unknown, pageOver = {}, blocks?: Record<string, unknown>) {
  return documentXml({
    ...await build(<Document id="d" title="D">{node as never}</Document>),
    style: { ...cleanMinimalDocumentStyle, page: columnPage(pageOver), ...(blocks ? { blocks } : {}) },
  } as never);
}

const hero = <Image id="m" src={pngWide} alt="M" width="full" />;

test("full keeps its keyword in the model", async () => {
  const doc = await build(<Document id="d" title="D">{hero}</Document>);
  assertEquals((doc.nodes[0] as ImageNode).width, "full");
});

test("full is the text column of the page as turned", async () => {
  assertStringIncludes(await styled(hero), '<wp:extent cx="6124575" cy="3057525"/>');
  assertStringIncludes(await styled(hero, { orientation: "landscape" }), '<wp:extent cx="9248775" cy="4629150"/>');
  assertStringIncludes(await styled(hero, { size: "LETTER" }), '<wp:extent cx="6334125" cy="3162300"/>');
});

test("full with a height draws exactly that box", async () => {
  assertStringIncludes(
    await styled(<Image id="m" src={pngWide} alt="M" width="full" height={100} />),
    '<wp:extent cx="6124575" cy="1266825"/>',
  );
});

test("full in a cell is the inside of the cell", async () => {
  const xml = await styled(
    <Table id="t" columns={[{ width: "auto" }, { width: "auto" }]}>
      <Row><Cell>{hero}</Cell><Cell><Paragraph id="p">x</Paragraph></Cell></Row>
    </Table>,
  );

  assertStringIncludes(elementOf(elementOf(xml, "w:tbl", "<w:drawing>"), "w:p", "<w:drawing>"), '<wp:extent cx="2857500" cy="1428750"/>');
});

test("a card's picture is the column less its padding", async () => {
  const xml = await styled(
    <Image id="m" variant="card" src={pngWide} alt="M" width="full" />,
    {},
    { card: { border: "D1D5DB", paddingPt: 4 } },
  );

  assertStringIncludes(xml, '<wp:extent cx="6019800"');
  assertStringIncludes(elementOf(xml, "w:tbl", "<w:drawing>"), '<w:tblW w:type="dxa" w:w="9638"/>');
});

// ---------------------------------------------------------------------------
// A picture that links somewhere
// ---------------------------------------------------------------------------

const PAY = "https://pay.example.com/inv/123";
const HYPERLINK_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

async function linked(href: string | undefined) {
  const doc = await build(
    <Document id="d" title="D">
      <Image id="qr" src={pngPixel} href={href} alt="Pay" width={8} height={8} />
    </Document>,
  );

  return { doc, xml: await documentXml(doc) };
}

test("a picture's address lands on its node, and is absent when not given", async () => {
  const withLink = (await linked(PAY)).doc.nodes[0] as ImageNode;
  const without = (await linked(undefined)).doc.nodes[0] as ImageNode;

  assertEquals(withLink.href, PAY);
  assertEquals("href" in without, false);
});

test("a linked picture is linked where Word looks, and the address is external", async () => {
  // One pack for both parts: the relationship id is minted afresh on every pack.
  const bytes = new Uint8Array(await (await createDocxBlob((await linked(PAY)).doc)).arrayBuffer());
  const read = (part: string) => new TextDecoder().decode(entryOf(bytes, part));
  const xml = read("word/document.xml");
  const id = /<w:hyperlink w:history="1" r:id="([^"]+)"><w:r><w:drawing>/.exec(xml)?.[1] ?? "";

  assertEquals(id !== "", true);
  assertStringIncludes(xml, `r:id="${id}"/></wp:docPr>`);
  assertStringIncludes(xml, `<a:hlinkClick r:id="${id}"/></pic:cNvPr>`);
  assertStringIncludes(
    read("word/_rels/document.xml.rels"),
    `Id="${id}" Type="${HYPERLINK_REL}" Target="${PAY}" TargetMode="External"`,
  );
});

test("a picture with no address packs as it always did", async () => {
  const { xml } = await linked(undefined);

  assertStringIncludes(xml, "<w:drawing>");
  assertEquals(xml.includes("<w:hyperlink"), false);
  assertEquals(xml.includes("hlinkClick"), false);
});

test("a refused scheme or an unsettled token draws the picture alone", async () => {
  for (const href of ["javascript:alert(1)", "{{data.payUrl}}"]) {
    const { xml } = await linked(href);

    assertStringIncludes(xml, "<w:drawing>");
    assertEquals(xml.includes("<w:hyperlink"), false);
    assertEquals(xml.includes("hlinkClick"), false);
  }
});

test("a picture among a paragraph's words links in place", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Paragraph id="p">
        Scan to pay <Image id="qr" src={pngPixel} href={PAY} width={8} height={8} />
      </Paragraph>
    </Document>,
  );

  assertEquals(
    /Scan to pay <\/w:t><\/w:r><w:hyperlink [^>]*><w:r><w:drawing>/.test(await documentXml(doc)),
    true,
  );
});

test("a picture in a table cell links", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Table id="t" columns={[{ width: "auto" }]}>
        <Row>
          <Cell>
            <Image id="qr" src={pngPixel} href={PAY} width={8} height={8} />
          </Cell>
        </Row>
      </Table>
    </Document>,
  );
  const cell = /<w:tc>.*<\/w:tc>/s.exec(await documentXml(doc))?.[0] ?? "";

  assertEquals(/<w:hyperlink[^]*<w:drawing>/.test(cell), true);
});

test("a picture in the header links through the header's own relationships", async () => {
  const doc = await build(
    <Document id="d" title="D" header={<Image id="logo" src={pngPixel} href={PAY} width={8} height={8} />}>
      <Paragraph id="p">Body</Paragraph>
    </Document>,
  );

  assertStringIncludes(await partXml(doc, "word/header1.xml"), "<w:hyperlink");
  assertStringIncludes(await partXml(doc, "word/_rels/header1.xml.rels"), `Target="${PAY}" TargetMode="External"`);
});

test("a picture that cannot be packed still sends the reader on", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Image id="qr" src="https://example.com/qr.png" alt="Pay now" href={PAY} />
    </Document>,
  );

  assertEquals(
    /<w:hyperlink [^>]*><w:r><w:t[^>]*>\[image: Pay now\]<\/w:t>/.test(await documentXml(doc)),
    true,
  );
});

test("a published picture keeps its address token, and settling fills it in", async () => {
  const Pay = () => {
    const [state] = useState((data: { payUrl: string }) => data);

    return <Image id="qr" src={pngPixel} href={state.payUrl} />;
  };
  const pay = template<{ payUrl: string }>(
    <Document id="d" title="D">
      <Pay />
    </Document>,
  );
  const data = { payUrl: "https://example.com/pay/42" };
  const published = await buildDocument(pay, createPublishData() as typeof data, {
    branchMode: "publish",
    deriverMode: "preserve",
    dynamicMode: "resolve",
    aiClient: { generateParagraph: () => "" },
  });
  const state: RuntimeState = {
    ctx: { ...data },
    derived: {},
    dataProvider: new InMemoryDataProvider(data),
    aiClient: new EchoAiClient(),
  };

  assertEquals((published.nodes[0] as ImageNode).href, "{{data.payUrl}}");

  const settled = await resolveDocument(published, state, { derivers: createDeriverRegistry() });
  const built = await buildDocument(pay, data, { branchMode: "decide", dynamicMode: "placeholder" });

  assertEquals((settled.nodes[0] as ImageNode).href, data.payUrl);
  assertEquals(JSON.stringify(settled.nodes[0]), JSON.stringify(built.nodes[0]));
});

test("a generated picture's address is settled too", async () => {
  const Pay = () => {
    const [state] = useState((data: { payUrl: string }) => data);

    return <Image id="qr" generalPrompt="A QR code." href={state.payUrl} />;
  };
  const data = { payUrl: "https://example.com/pay/42" };
  const published = await buildDocument(
    template<{ payUrl: string }>(
      <Document id="d" title="D">
        <Pay />
      </Document>,
    ),
    createPublishData() as typeof data,
    {
      branchMode: "publish",
      deriverMode: "preserve",
      dynamicMode: "placeholder",
      aiClient: { generateParagraph: () => "" },
    },
  );
  const state: RuntimeState = {
    ctx: { ...data },
    derived: {},
    dataProvider: new InMemoryDataProvider(data),
    aiClient: { ...new EchoAiClient(), generateParagraph: () => "", generateImage: () => ({ path: pngPixel }) },
  };

  assertEquals((published.nodes[0] as ImageNode).href, "{{data.payUrl}}");
  assertEquals(
    ((await resolveDocument(published, state, { derivers: createDeriverRegistry() })).nodes[0] as ImageNode).href,
    data.payUrl,
  );
});

test("the preview keeps a linked picture's frame inside its link", () => {
  const dom = new JSDOM('<div id="page"><p></p></div>');
  const page = dom.window.document.getElementById("page") as Element;
  // Built node by node: the HTML parser would end the <p> at a <div>, which
  // docx-preview's DOM calls do not.
  const link = page.ownerDocument.createElement("a");
  const frame = page.ownerDocument.createElement("div");

  link.setAttribute("href", PAY);
  frame.setAttribute("style", "display: inline-block; width: 10pt; height: 10pt;");
  frame.append(page.ownerDocument.createElement("img"));
  link.append(frame);
  page.querySelector("p")?.append(link);

  settleDocxPreview(page);

  assertEquals(page.querySelector(`p > a[href="${PAY}"] > span > img`) !== null, true);
  assertEquals(page.querySelector("p div"), null);
});

const statusStyle = {
  ...cleanMinimalDocumentStyle,
  blocks: { ...cleanMinimalDocumentStyle.blocks, status: { color: "1F7A3D" } },
};

const paidCell = () =>
  build(
    <Document id="d" title="D">
      <Table id="t" columns={[{ width: "auto" }]}>
        <Row>
          <Cell variant="status">
            <Paragraph id="p">
              Paid <Image id="i" src={pngPixel} width={8} height={8} /> on 3 March
            </Paragraph>
          </Cell>
        </Row>
      </Table>
    </Document>,
  );

test("a picture mid-sentence in a cell sits where it was written", async () => {
  const xml = await documentXml({ ...(await paidCell()), style: statusStyle });
  const p = paragraphOf(xml, "Paid ");

  assertEquals(p.indexOf("Paid ") < p.indexOf("<w:drawing>"), true);
  assertEquals(p.indexOf("<w:drawing>") < p.indexOf(" on 3 March"), true);
  assertEquals((p.match(/<w:color w:val="1F7A3D"\/>/g) ?? []).length, 2);
});

test("a picture mid-sentence in a shape sits where it was written", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Shape id="s">
          <Paragraph id="p">Look for the <Image id="i" src={pngPixel} width={8} height={8} /> mark</Paragraph>
        </Shape>
      </Document>,
    ),
  );
  const p = paragraphOf(xml, "Look for the ");

  assertEquals(p.indexOf("Look for the ") < p.indexOf("<w:drawing>"), true);
  assertEquals(p.indexOf("<w:drawing>") < p.indexOf(" mark"), true);
});

test("a cell picture at the start or past the end keeps its place", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="t" columns={[{ width: "auto" }]}>
          <Row>
            <Cell>
              <Paragraph id="p">
                <Image id="a" src={pngPixel} width={8} height={8} />Due <Image id="b" src={pngPixel} width={8} height={8} />
              </Paragraph>
            </Cell>
          </Row>
        </Table>
      </Document>,
    ),
  );
  const p = paragraphOf(xml, "Due ");

  assertEquals(/<w:drawing>[\s\S]*Due <\/w:t><\/w:r>[\s\S]*<w:drawing>/.test(p), true);
  assertEquals((p.match(/<w:drawing>/g) ?? []).length, 2);
  assertEquals((p.match(/<w:t[ >]/g) ?? []).length, 1);
});

test("a cell paragraph with no picture is still one run", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="t" columns={[{ width: "auto" }]}>
          <Row>
            <Cell>
              <Paragraph id="p">Just words.</Paragraph>
            </Cell>
          </Row>
        </Table>
      </Document>,
    ),
  );
  const p = paragraphOf(xml, "Just words.");

  assertEquals((p.match(/<w:r>/g) ?? []).length, 1);
  assertEquals(p.includes("<w:drawing>"), false);
});

test("a cell paragraph's picture offset is recorded in the model", async () => {
  // The paragraph sits inside table, row and cell nodes; JSON.parse's reviver visits all of them.
  let node = {} as ParagraphNode;
  JSON.parse(JSON.stringify(await paidCell()), (_key, value) => {
    if (value?.kind === "paragraph" && value.id === "p") node = value;
    return value;
  });

  assertEquals(node.inlineImages?.length, 1);
  assertEquals(node.inlineImages?.[0].at, 5);
  assertEquals(node.text, "Paid  on 3 March");
});
