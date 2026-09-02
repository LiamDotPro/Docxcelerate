import { test } from "node:test";
import { assertEquals, assertStringIncludes } from "./assert.ts";
import { buildDocument } from "docxcelerate";
import { documentXml } from "./docx.ts";
import type { DocumentModel, DocumentStyle } from "docxcelerate";
import { Document, Paragraph, template } from "docxcelerate/template";

/**
 * What a paragraph can be told to do, and what Word is handed when it is.
 *
 * Everything here is checked against the packed file rather than against the
 * model, because a property that exists in the model and nowhere in the `.docx`
 * is a property that does nothing — which is how `radiusPt` came to be a
 * setting nobody could see the effect of. The same claims are measured against
 * Word itself and against the preview by the conformance suite; these are the
 * half that can run on any machine, and they are what a regression trips
 * first.
 *
 * The numbers: Word counts spacing and indents in twentieths of a point, so a
 * point is 20 and a millimetre is 56.6929… twips. 10pt is 200, 10mm is 567.
 */

const style: DocumentStyle = {
  preset: "test",
  page: {
    size: "A4",
    orientation: "portrait",
    margins: { topMm: 20, rightMm: 20, bottomMm: 20, leftMm: 20 },
  },
  typography: {
    bodyFont: "Aptos",
    headingFont: "Aptos",
    bodySizePt: 11,
    bodyLineHeight: 1.4,
    color: "111827",
  },
  paragraph: { spacingAfterPt: 10 },
  title: { fontSizePt: 20, weight: "bold", spacingBeforePt: 0, spacingAfterPt: 18 },
  sectionHeading: { fontSizePt: 12, weight: "bold", spacingBeforePt: 16, spacingAfterPt: 7 },
  showTitle: false,
  blocks: {
    /** A theme deciding what one of its named blocks looks like. */
    standfirst: { align: "center" },
    /** Inset from both margins, the way a pulled quotation sits. */
    quote: { indentMm: 10, indentRightMm: 10 },
    /** A book's paragraph mark: the first line, and nothing else. */
    firstLine: { firstLineIndentMm: 8 },
    /** A definition whose later lines clear its term. */
    hanging: { indentMm: 12, hangingIndentMm: 12 },
    /** Both indents at once, which Word cannot do and the packer must resolve. */
    confused: { firstLineIndentMm: 6, hangingIndentMm: 9 },
    /** Set apart from whatever ran before it. */
    interlude: { spacingBeforePt: 18, spacingAfterPt: 6 },
    /** Never the last thing on a page. */
    stayWithNext: { keepWithNext: true },
    /** One thing, not two halves. */
    stayWhole: { keepLines: true },
    /** A contents line: text, dots, a page number on the right margin. */
    contentsLine: {
      tabStopsMm: [{ at: 170, align: "right", leader: "dot" }],
    },
    /** A signature block: a name, then a date halfway across. */
    signature: { tabStopsMm: [{ at: 90 }] },
    /** A bleed still wins over an indent — the two are different intentions. */
    bleedingQuote: { bleed: true, indentMm: 10 },
  },
};

function build(node: unknown): Promise<DocumentModel> {
  return buildDocument(
    template<Record<string, never>>(node as never),
    {},
    { branchMode: "decide", dynamicMode: "placeholder" },
  ).then((doc) => ({ ...doc, style }));
}

/** The `<w:p>` whose text contains this anchor. */
function paragraphOf(xml: string, anchor: string): string {
  const paragraphs = [...xml.matchAll(/<w:p>[\s\S]*?<\/w:p>/g)].map((match) => match[0]);
  const found = paragraphs.find((paragraph) => paragraph.includes(anchor));

  if (found === undefined) {
    throw new Error(`no paragraph containing ${JSON.stringify(anchor)}`);
  }

  return found;
}

// ---------------------------------------------------------------------------
// Alignment
// ---------------------------------------------------------------------------

test("a paragraph says how its lines sit, and Word is told in its own word", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a">Left as written.</Paragraph>
        <Paragraph id="b" align="center">Centred.</Paragraph>
        <Paragraph id="c" align="right">Ranged right.</Paragraph>
        <Paragraph id="d" align="justify">Justified.</Paragraph>
      </Document>,
    ),
  );

  assertStringIncludes(paragraphOf(xml, "Centred."), '<w:jc w:val="center"/>');
  assertStringIncludes(paragraphOf(xml, "Ranged right."), '<w:jc w:val="right"/>');
  // Word calls justification "both", meaning both edges flush. The document
  // says the word people say and the translation stops at the packer.
  assertStringIncludes(paragraphOf(xml, "Justified."), '<w:jc w:val="both"/>');
});

test("a paragraph that never said how to sit writes no alignment at all", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a">Left as written.</Paragraph>
      </Document>,
    ),
  );

  // Not `w:jc w:val="left"`. Writing an explicit left would override whatever
  // style the paragraph sits inside, which is a different statement from
  // having no opinion.
  assertEquals(paragraphOf(xml, "Left as written.").includes("<w:jc"), false);
});

test("a theme can align a named block, without the node saying anything", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a" variant="standfirst">Centred by the block.</Paragraph>
      </Document>,
    ),
  );

  assertStringIncludes(paragraphOf(xml, "Centred by the block."), '<w:jc w:val="center"/>');
});

test("a node's alignment beats its block's, the way a cell beats its column", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a" variant="standfirst" align="right">Ranged, not centred.</Paragraph>
      </Document>,
    ),
  );

  assertStringIncludes(paragraphOf(xml, "Ranged, not centred."), '<w:jc w:val="right"/>');
  assertEquals(paragraphOf(xml, "Ranged, not centred.").includes('w:val="center"'), false);
});

// ---------------------------------------------------------------------------
// Indentation
// ---------------------------------------------------------------------------

test("a block inset from the margins is inset in the file, in twips", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a" variant="quote">Inset from both margins.</Paragraph>
      </Document>,
    ),
  );

  // 10mm is 567 twips. Positive, because this sits inside the text column —
  // reaching outside it is what `bleed` is for, and it is negative.
  assertStringIncludes(paragraphOf(xml, "Inset from both margins."), 'w:left="567"');
  assertStringIncludes(paragraphOf(xml, "Inset from both margins."), 'w:right="567"');
});

test("a first-line indent moves the first line and leaves the block where it was", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a" variant="firstLine">Indented on the first line.</Paragraph>
      </Document>,
    ),
  );

  const paragraph = paragraphOf(xml, "Indented on the first line.");

  assertStringIncludes(paragraph, 'w:firstLine="454"');
  assertEquals(paragraph.includes("w:left="), false);
});

test("a hanging block pulls its first line back from the rest", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a" variant="hanging">Hanging.</Paragraph>
      </Document>,
    ),
  );

  const paragraph = paragraphOf(xml, "Hanging.");

  assertStringIncludes(paragraph, 'w:left="680"');
  assertStringIncludes(paragraph, 'w:hanging="680"');
});

test("a hang and a first-line indent are one attribute, and the hang wins", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a" variant="confused">Both at once.</Paragraph>
      </Document>,
    ),
  );

  const paragraph = paragraphOf(xml, "Both at once.");

  // Word writes one or the other, never both: they are the same attribute
  // pulling opposite ways. The hang is the structural one, so it wins — and
  // the file must not carry a first line that contradicts it.
  assertStringIncludes(paragraph, 'w:hanging="510"');
  assertEquals(paragraph.includes("w:firstLine="), false);
});

test("a bleed is still a bleed, whatever indent the block also names", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a" variant="bleedingQuote">Out past the margin.</Paragraph>
      </Document>,
    ),
  );

  // Negative, by the width of the margin: reaching past the sheet's edge and
  // sitting inside the column are different intentions, and a block that
  // states both gets the one it named first.
  assertStringIncludes(paragraphOf(xml, "Out past the margin."), 'w:left="-1134"');
});

test("a paragraph with no indent of its own writes no indent element", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a">Flush to both margins.</Paragraph>
      </Document>,
    ),
  );

  assertEquals(paragraphOf(xml, "Flush to both margins.").includes("<w:ind"), false);
});

// ---------------------------------------------------------------------------
// Spacing
// ---------------------------------------------------------------------------

test("a block can leave space above itself, not only below", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a">Ordinary prose.</Paragraph>
        <Paragraph id="b" variant="interlude">Set apart.</Paragraph>
      </Document>,
    ),
  );

  const paragraph = paragraphOf(xml, "Set apart.");

  assertStringIncludes(paragraph, 'w:before="360"');
  assertStringIncludes(paragraph, 'w:after="120"');
});

test("a paragraph that says nothing about space above writes nothing", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a">Ordinary prose.</Paragraph>
      </Document>,
    ),
  );

  const paragraph = paragraphOf(xml, "Ordinary prose.");

  assertEquals(paragraph.includes("w:before="), false);
  // The document's own space-after still applies, because that is a default
  // rather than a silence.
  assertStringIncludes(paragraph, 'w:after="200"');
});

// ---------------------------------------------------------------------------
// Keeps
// ---------------------------------------------------------------------------

test("a block can refuse to be the last thing on a page", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a" variant="stayWithNext">A heading.</Paragraph>
        <Paragraph id="b" variant="stayWhole">An address block.</Paragraph>
      </Document>,
    ),
  );

  assertStringIncludes(paragraphOf(xml, "A heading."), "<w:keepNext/>");
  assertStringIncludes(paragraphOf(xml, "An address block."), "<w:keepLines/>");
});

test("a paragraph with no opinion about breaks writes neither keep", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a">Ordinary prose.</Paragraph>
      </Document>,
    ),
  );

  const paragraph = paragraphOf(xml, "Ordinary prose.");

  // Not `<w:keepNext w:val="0"/>`. Telling Word to allow a break is a
  // statement, and it would override a style that had asked for one.
  assertEquals(paragraph.includes("keepNext"), false);
  assertEquals(paragraph.includes("keepLines"), false);
});

// ---------------------------------------------------------------------------
// Tab stops
// ---------------------------------------------------------------------------

test("a block's tab stops are written where a ruler would put them", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a" variant="contentsLine">{"What the agreement covers\t3"}</Paragraph>
      </Document>,
    ),
  );

  const paragraph = paragraphOf(xml, "What the agreement covers");

  // 170mm from the left margin, which on an A4 page with 20mm margins is the
  // right margin exactly.
  assertStringIncludes(paragraph, '<w:tab w:val="right" w:pos="9638" w:leader="dot"/>');
});

test("a tab in the text is written as a tab, not as whitespace", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a" variant="contentsLine">{"What the agreement covers\t3"}</Paragraph>
      </Document>,
    ),
  );

  const paragraph = paragraphOf(xml, "What the agreement covers");

  // The run has to carry `<w:tab/>`. A literal tab character left inside
  // `<w:t>` is whitespace: Word is forgiving enough to draw it, which is worse
  // than if it were not, because the file then looks right in Word and
  // collapses to one space everywhere else — including in the preview.
  assertStringIncludes(paragraph, "<w:tab/>");
  assertEquals(/<w:t[^>]*>[^<]*\t/.test(paragraph), false);
});

test("a stop with nothing said about it is a plain left stop", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a" variant="signature">{"Signed for the supplier\tDate"}</Paragraph>
      </Document>,
    ),
  );

  assertStringIncludes(
    paragraphOf(xml, "Signed for the supplier"),
    '<w:tab w:val="left" w:pos="5102" w:leader="none"/>',
  );
});

test("a paragraph with no stops writes no tabs element", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a">Ordinary prose.</Paragraph>
      </Document>,
    ),
  );

  assertEquals(paragraphOf(xml, "Ordinary prose.").includes("<w:tabs>"), false);
});

test("the text either side of a tab survives it", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a" variant="signature">{"Signed for the supplier\tDate"}</Paragraph>
      </Document>,
    ),
  );

  const paragraph = paragraphOf(xml, "Signed for the supplier");

  assertStringIncludes(paragraph, "Signed for the supplier");
  assertStringIncludes(paragraph, "Date");
});
