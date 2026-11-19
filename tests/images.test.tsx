import { test } from "node:test";
import { assertEquals, assertStringIncludes } from "./assert.ts";
import { buildDocument, imageSourceOf, isSvg, rasterTypeOf } from "docxcelerate";
import { createDocxBlob, createDocxDocument } from "docxcelerate/docx";

import {
  Cell,
  Document,
  Image,
  Paragraph,
  Row,
  Table,
  createPublishData,
  template,
  useState,
} from "docxcelerate/template";
import { createDeriverRegistry, EchoAiClient, InMemoryDataProvider, resolveDocument } from "docxcelerate";
import type { ImageNode, RuntimeState } from "docxcelerate";
import { settleDocxPreview } from "docxcelerate/preview";
import { JSDOM } from "jsdom";
import { documentXml, entryOf, partXml } from "./docx.ts";

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
