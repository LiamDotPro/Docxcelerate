import { test } from "node:test";
import { assertEquals, assertStringIncludes } from "./assert.ts";
import { buildDocument } from "docxcelerate";
import { documentXml } from "./docx.ts";
import type { DocumentModel, DocumentStyle } from "docxcelerate";
import {
  Cell,
  Document,
  Image,
  PageBreak,
  PageNumber,
  Paragraph,
  Row,
  Table,
  template,
} from "docxcelerate/template";

/**
 * What Word is actually handed.
 *
 * A `.docx` is a zip of XML, so the way to know a fill or a border survived
 * packing is to open the file and look. Without this, a style could silently
 * mean one thing where it is authored and nothing at all in the format the
 * framework exists to produce — which is how `radiusPt` came to be a property
 * that did nothing here.
 */

const style: DocumentStyle = {
  preset: "test",
  page: {
    size: "A4",
    orientation: "portrait",
    margins: { topMm: 16, rightMm: 16, bottomMm: 16, leftMm: 16 },
  },
  typography: {
    bodyFont: "Aptos",
    headingFont: "Aptos",
    bodySizePt: 10,
    bodyLineHeight: 1.4,
    color: "1C2340",
  },
  palette: {
    heading: "2C3D8F",
    accent: "2C3D8F",
    muted: "5A6482",
    rule: "D9DDEB",
    page: "FFFFFF",
  },
  paragraph: { spacingAfterPt: 6 },
  title: { fontSizePt: 20, weight: "bold", spacingBeforePt: 0, spacingAfterPt: 10 },
  sectionHeading: { fontSizePt: 8, weight: "bold", spacingBeforePt: 10, spacingAfterPt: 4 },
  blocks: {
    band: {
      fill: "F4F6FD",
      bleed: true,
      border: "E3E7F5",
      borderSides: ["bottom"],
      paddingPt: 10,
    },
    totalRow: {
      fill: "1E2A66",
      color: "FFFFFF",
      weight: "bold",
      fontSizePt: 11,
    },
    badge: {
      fill: "FBF0DC",
      border: "E5C78A",
      color: "8A5A06",
      paddingPt: 5,
      fontSizePt: 7,
      weight: "bold",
      transform: "uppercase",
      letterSpacingEm: 0.1,
    },
  },
};

/** A one-pixel PNG, so a picture test has bytes to embed. */
const PIXEL = "data:image/png;base64," +
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

function build(node: unknown): Promise<DocumentModel> {
  return buildDocument(
    template<Record<string, never>>(node as never),
    {},
    { branchMode: "decide", dynamicMode: "placeholder" },
  ).then((doc) => ({ ...doc, style }));
}

// ---------------------------------------------------------------------------
// A block means the same thing in both outputs
// ---------------------------------------------------------------------------

test("a block's fill is shading in the Word file, not just colour on screen", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a" variant="band">Dates.</Paragraph>
      </Document>,
    ),
  );

  assertStringIncludes(xml, 'w:fill="F4F6FD"');
});

test("a block's border is drawn, on the edges it named", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a" variant="band">Dates.</Paragraph>
      </Document>,
    ),
  );

  assertStringIncludes(xml, "<w:pBdr>");
  assertStringIncludes(xml, 'w:color="E3E7F5"');
  assertStringIncludes(xml, "<w:bottom");
  // Only the bottom was asked for, so no other edge is drawn.
  assertEquals(/<w:pBdr>(?:(?!<\/w:pBdr>).)*<w:top /s.test(xml), false);
});

test("a bleeding block reaches past the margin, by indenting negatively", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a" variant="band">Dates.</Paragraph>
      </Document>,
    ),
  );

  // 16mm of margin, so the block starts 16mm to the left of it.
  assertStringIncludes(xml, 'w:left="-907"');
});

test("letter spacing survives, converted from ems to what Word counts in", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="t" columns={[{ width: "auto" }]}>
          <Row>
            <Cell id="c" variant="badge">Awaiting</Cell>
          </Row>
        </Table>
      </Document>,
    ),
  );

  // 0.1em of a 7pt face is 0.7pt, and Word counts in twentieths of a point.
  assertStringIncludes(xml, '<w:spacing w:val="14"/>');
});

test("a cell's variant sets its text, not only its background", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="t" columns={[{ width: "auto" }]}>
          <Row>
            <Cell id="c" variant="badge">Awaiting</Cell>
          </Row>
        </Table>
      </Document>,
    ),
  );

  assertStringIncludes(xml, 'w:fill="FBF0DC"');
  assertStringIncludes(xml, 'w:val="8A5A06"');
  assertStringIncludes(xml, "<w:caps/>");
  assertStringIncludes(xml, "<w:tcBorders>");
  assertStringIncludes(xml, "<w:tcMar>");
});

test("a variant the theme has never heard of packs as an ordinary block", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a" variant="nonesuch">Text.</Paragraph>
      </Document>,
    ),
  );

  assertEquals(xml.includes("<w:pBdr>"), false);
  assertStringIncludes(xml, "Text.");
});

// ---------------------------------------------------------------------------
// The furniture
// ---------------------------------------------------------------------------

test("a page break is Word's own break, not a run of empty paragraphs", async () => {
  const broken = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a">One.</Paragraph>
        <PageBreak id="turn" />
        <Paragraph id="b">Two.</Paragraph>
      </Document>,
    ),
  );
  const unbroken = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a">One.</Paragraph>
        <Paragraph id="b">Two.</Paragraph>
      </Document>,
    ),
  );

  assertStringIncludes(broken, 'w:type="page"');
  assertEquals(unbroken.includes('w:type="page"'), false);
});

test("a page number is a field Word recounts, not a digit written in", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D" footer={<PageNumber id="n" />}>
        <Paragraph id="a">One.</Paragraph>
      </Document>,
    ),
  );

  // The number lives in the footer part rather than the body, which is itself
  // the point — it is furniture, and Word fills it in per page.
  assertEquals(xml.includes("PAGE"), false);
});

test("a title the document prints itself is not printed again", async () => {
  const withTitle = await documentXml(
    await build(
      <Document id="d" title="Invoice">
        <Paragraph id="a">One.</Paragraph>
      </Document>,
    ),
  );
  const withoutTitle = await documentXml(
    await buildDocument(
      template<Record<string, never>>(
        (
          <Document id="d" title="Invoice">
            <Paragraph id="a">One.</Paragraph>
          </Document>
        ) as never,
      ),
      {},
      { branchMode: "decide", dynamicMode: "placeholder" },
    ).then((doc) => ({ ...doc, style: { ...style, showTitle: false } })),
  );

  assertStringIncludes(withTitle, "Invoice");
  assertEquals(withoutTitle.includes("Invoice"), false);
});
// ---------------------------------------------------------------------------
// What a table is drawn with
// ---------------------------------------------------------------------------

test("a table draws the theme's rules, not Word's default grid", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="t" columns={[{ width: "auto" }]}>
          <Row header>
            <Cell id="h">Description</Cell>
          </Row>
          <Row>
            <Cell id="c">A line.</Cell>
          </Row>
        </Table>
      </Document>,
    ),
  );

  // The library's own default is a black box around every cell. Left alone it
  // means every document ships a grid nobody asked for and no preview shows.
  assertEquals(/<w:tblBorders>(?:(?!<\/w:tblBorders>).)*w:val="single"/s.test(xml), false);
  // What separates one row from the next is the palette's rule, under the
  // body cells only — a heading has its fill to set it apart.
  assertStringIncludes(xml, 'w:color="D9DDEB"');
  assertEquals(xml.match(/w:color="D9DDEB"/g)?.length, 1);
});

test("a header row the theme names keeps its own fill, not the accent", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="t" columns={[{ width: "auto" }, { width: "auto" }]}>
          <Row header variant="totalRow">
            <Cell id="empty"></Cell>
            <Cell id="total">Total due</Cell>
          </Row>
        </Table>
      </Document>,
    ),
  );

  // Every cell in the row, including the empty one beside the words: a bar in
  // two colours is what happens when the default shows through the gap.
  assertEquals(xml.match(/w:fill="1E2A66"/g)?.length, 2);
  assertEquals(xml.includes('w:fill="2C3D8F"'), false);
});

test("a heading row is set in small tracked capitals, as the screen sets it", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="t" columns={[{ width: "auto" }]}>
          <Row header>
            <Cell id="h">Description</Cell>
          </Row>
        </Table>
      </Document>,
    ),
  );

  assertStringIncludes(xml, "<w:caps/>");
  assertStringIncludes(xml, 'w:val="FFFFFF"');
  // 0.72 of a 10pt body, in half-points.
  assertStringIncludes(xml, '<w:sz w:val="14"/>');
});

test("a picture in a cell is not laid out as a paragraph of prose", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="t" columns={[{ width: "auto" }]}>
          <Row>
            <Cell id="c">
              <Image id="mark" src={PIXEL} alt="" width={8} height={8} />
            </Cell>
          </Row>
        </Table>
      </Document>,
    ),
  );

  // The body's line and space-after belong to prose. On a picture they drop a
  // letterhead's mark below the name beside it and make a one-line footer bar
  // two lines deep.
  assertEquals(
    /<w:spacing w:after="120"[^>]*\/>(?:(?!<\/w:p>).)*<w:drawing>/s.test(xml),
    false,
  );
  assertStringIncludes(xml, "<w:drawing>");
});
