import { test } from "node:test";
import { assertEquals, assertStringIncludes } from "./assert.ts";
import { buildDocument } from "docxcelerate";
import { createDocxDocument } from "docxcelerate/docx";
import { documentXml, partNames, partXml } from "./docx.ts";
import type { DocumentModel, DocumentStyle } from "docxcelerate";
import {
  Cell,
  Document,
  PageBreak,
  PageNumber,
  Paragraph,
  Row,
  Table,
  template,
} from "docxcelerate/template";

/**
 * Page furniture: the parts of a document that are not its text.
 *
 * A break, a running header, a page number and a block style are all things a
 * build cannot settle on its own — how many pages a document runs to depends on
 * what an engine writes into it, and what a `band` looks like depends on the
 * theme it is read under. What the build can do is say which is which, and that
 * is what these check, in the model and then in the file it packs into.
 */

const styleWithBlocks: DocumentStyle = {
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
  paragraph: { spacingAfterPt: 6 },
  title: { fontSizePt: 20, weight: "bold", spacingBeforePt: 0, spacingAfterPt: 10 },
  sectionHeading: { fontSizePt: 8, weight: "bold", spacingBeforePt: 10, spacingAfterPt: 4 },
  blocks: {
    band: { fill: "F4F6FD", paddingPt: 10 },
    badge: { fill: "FBF0DC", border: "E5C78A", color: "8A5A06", transform: "uppercase" },
  },
};

function build(node: unknown, style?: DocumentStyle): Promise<DocumentModel> {
  return buildDocument(
    template<Record<string, never>>(node as never),
    {},
    { branchMode: "decide", dynamicMode: "placeholder" },
  ).then((doc) => (style ? { ...doc, style } : doc)) as Promise<DocumentModel>;
}

/** How many breaks the packed document turns a page on. */
function breaksIn(xml: string): number {
  return (xml.match(/w:type="page"/g) ?? []).length;
}

// ---------------------------------------------------------------------------
// Where a page ends
// ---------------------------------------------------------------------------

test("a document with no break runs on rather than turning a page", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a">One.</Paragraph>
        <Paragraph id="b">Two.</Paragraph>
      </Document>,
    ),
  );

  assertEquals(breaksIn(xml), 0);
});

test("a break turns the page where the document said, and prints nothing itself", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="owed">What is owed.</Paragraph>
        <PageBreak id="turn" />
        <Paragraph id="pay">How to pay it.</Paragraph>
      </Document>,
    ),
  );

  assertEquals(breaksIn(xml), 1);
  // Between the two, which is the whole of what a break says.
  assertEquals(xml.indexOf("What is owed.") < xml.indexOf('w:type="page"'), true);
  assertEquals(xml.indexOf('w:type="page"') < xml.indexOf("How to pay it."), true);
  // And it is a break, not a run of empty paragraphs standing in for one.
  assertEquals(xml.includes("turn"), false);
});

// ---------------------------------------------------------------------------
// What repeats
// ---------------------------------------------------------------------------

test("running furniture is packed as Word's own, so it repeats on every page", async () => {
  const doc = await build(
    <Document
      id="d"
      title="D"
      header={<Paragraph id="head">INV-2026-0142</Paragraph>}
      footer={<Paragraph id="foot">Fernhill Systems Ltd</Paragraph>}
    >
      <Paragraph id="a">One.</Paragraph>
      <PageBreak id="turn" />
      <Paragraph id="b">Two.</Paragraph>
    </Document>,
  );

  const names = await partNames(doc);

  // A header part and a footer part, written once. Word draws them on each
  // page it repaginates to — which is why the count of pages is never a number
  // this framework has to guess at.
  assertEquals(names.includes("word/header1.xml"), true);
  assertEquals(names.includes("word/footer1.xml"), true);
  assertStringIncludes(await partXml(doc, "word/header1.xml"), "INV-2026-0142");
  assertStringIncludes(await partXml(doc, "word/footer1.xml"), "Fernhill Systems Ltd");
  // Not folded into the body, where it would print once and in the wrong place.
  assertEquals((await documentXml(doc)).includes("INV-2026-0142"), false);
});

test("furniture is kept apart from the body, not folded into it", async () => {
  const doc = await build(
    <Document id="d" title="D" header={<Paragraph id="head">Head.</Paragraph>}>
      <Paragraph id="a">Body.</Paragraph>
    </Document>,
  );

  assertEquals(doc.header?.length, 1);
  assertEquals(doc.nodes.map((node) => node.id), ["a"]);
});

test("a document with no furniture carries none, rather than carrying nothing", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Paragraph id="a">Body.</Paragraph>
    </Document>,
  );

  assertEquals(doc.header, undefined);
  assertEquals(doc.footer, undefined);
});

test("the title is the document's name, so it is written once and not per page", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="Invoice">
        <Paragraph id="a">One.</Paragraph>
        <PageBreak id="turn" />
        <Paragraph id="b">Two.</Paragraph>
      </Document>,
    ),
  );

  assertEquals((xml.match(/w:val="Title"/g) ?? []).length, 1);
});

test("a document that prints its own title is not given a second one", async () => {
  // An invoice whose letterhead carries the wordmark beside the reference. The
  // title is still the document's name for anything reading the model; it is
  // just not printed above a page that already says it.
  const doc = await build(
    <Document id="d" title="Invoice">
      <Paragraph id="a">One.</Paragraph>
    </Document>,
    { ...styleWithBlocks, showTitle: false },
  );

  assertEquals((await documentXml(doc)).includes('w:val="Title"'), false);
  assertEquals(doc.title, "Invoice");
  assertEquals(typeof createDocxDocument(doc), "object");
});

test("a document that says nothing about its title still gets one", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="Invoice">
        <Paragraph id="a">One.</Paragraph>
      </Document>,
      styleWithBlocks,
    ),
  );

  assertStringIncludes(xml, 'w:val="Title"');
});

// ---------------------------------------------------------------------------
// Counting the pages
// ---------------------------------------------------------------------------

test("a page number says which form it wants rather than carrying a digit", async () => {
  const doc = await build(
    <Document id="d" title="D" footer={<PageNumber id="n" separator=" of " />}>
      <Paragraph id="a">One.</Paragraph>
    </Document>,
  );

  const foot = doc.footer?.[0];

  assertEquals(foot?.kind, "pageNumber");
  assertEquals(foot?.kind === "pageNumber" ? foot.separator : undefined, " of ");
});

test("the number is a field Word recounts, both halves of it", async () => {
  const doc = await build(
    <Document id="d" title="D" footer={<PageNumber id="n" />}>
      <Paragraph id="a">One.</Paragraph>
      <PageBreak id="turn" />
      <Paragraph id="b">Two.</Paragraph>
    </Document>,
  );

  const foot = await partXml(doc, "word/footer1.xml");

  // "1 of 2" written in at build time is a number that stops being true the
  // moment an engine writes a longer paragraph. These are instructions Word
  // re-evaluates whenever it repaginates.
  assertStringIncludes(foot, "PAGE");
  assertStringIncludes(foot, "NUMPAGES");
});

// ---------------------------------------------------------------------------
// A variant is a name, not an appearance
// ---------------------------------------------------------------------------

test("a variant travels on the node as the name the component wrote", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Paragraph id="a" variant="band">Dates.</Paragraph>
    </Document>,
  );

  assertEquals(doc.nodes[0].variant, "band");
});

test("the theme decides what the name looks like, and the node never says", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a" variant="band">Dates.</Paragraph>
      </Document>,
      styleWithBlocks,
    ),
  );

  assertStringIncludes(xml, 'w:fill="F4F6FD"');
});

test("a variant the theme has never heard of draws as an ordinary block", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a" variant="nonesuch">Text.</Paragraph>
      </Document>,
      styleWithBlocks,
    ),
  );

  assertEquals(xml.includes("<w:shd"), false);
  assertStringIncludes(xml, "Text.");
});

test("a block that does not bleed stays inside the text column", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a" variant="band">Dates.</Paragraph>
      </Document>,
      styleWithBlocks,
    ),
  );

  assertEquals(xml.includes('w:left="-907"'), false);
});

// ---------------------------------------------------------------------------
// A table is furniture too
// ---------------------------------------------------------------------------

test("a cell's variant reaches the cell, not the row it happens to sit in", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Table id="t" columns={[{ width: "auto" }, { width: "auto" }]}>
          <Row>
            <Cell id="plain">Due</Cell>
            <Cell id="status" variant="badge">Awaiting</Cell>
          </Row>
        </Table>
      </Document>,
      styleWithBlocks,
    ),
  );

  assertEquals((xml.match(/w:fill="FBF0DC"/g) ?? []).length, 1);
});
