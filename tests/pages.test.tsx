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
  Section,
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

/**
 * How many breaks the packed document turns a page on.
 *
 * A break between siblings travels as a break-carrying *style* named by
 * whatever follows it — a property of where the next thing starts, not a
 * paragraph of its own. The standalone `w:br w:type="page"` form survives only
 * where there is no following paragraph to carry the style.
 */
function breaksIn(xml: string): number {
  return (xml.match(/w:type="page"|w:val="PageBreakBefore/g) ?? []).length;
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
  // The break rides the next paragraph's style, so nothing — not even an empty
  // line — is left at the foot of the outgoing page.
  assertStringIncludes(xml, '<w:pStyle w:val="PageBreakBefore"/>');
  assertEquals(xml.includes('w:type="page"'), false);
  assertEquals(xml.indexOf("What is owed.") < xml.indexOf('w:val="PageBreakBefore"'), true);
  assertEquals(xml.indexOf('w:val="PageBreakBefore"') < xml.indexOf("How to pay it."), true);
  // And it is a break, not a run of empty paragraphs standing in for one.
  assertEquals(xml.includes("turn"), false);
});

test("a break into a table rides a hairline paragraph, not an empty line", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a">Before.</Paragraph>
        <PageBreak id="turn" />
        <Table id="t" columns={[{ width: "auto" }]}>
          <Row>
            <Cell id="c">After.</Cell>
          </Row>
        </Table>
      </Document>,
    ),
  );

  // A table carries no break of its own — Word has no break-before on one and
  // docx-preview splits only on top-level elements — so a paragraph goes in
  // front to carry it. It is the hairline separator wearing the break style,
  // so what lands is one point at the head of the new page rather than a full
  // empty line at the foot of the old one.
  assertEquals(breaksIn(xml), 1);
  assertEquals(xml.includes('w:type="page"'), false);
  assertStringIncludes(xml, '<w:pStyle w:val="PageBreakBefore"/>');
  assertStringIncludes(xml, '<w:spacing w:after="0" w:before="0" w:line="20" w:lineRule="exact"/>');
});

test("two tables in a row stay two tables, which Word needs saying explicitly", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Section id="one" title="One" showTitle={false}>
          <Table id="t1" columns={[{ width: "auto" }]}>
            <Row>
              <Cell id="c1">First.</Cell>
            </Row>
          </Table>
        </Section>
        <Section id="two" title="Two" showTitle={false}>
          <Table id="t2" columns={[{ width: "auto" }, { width: "auto" }]}>
            <Row>
              <Cell id="c2">Second.</Cell>
              <Cell id="c3">Also second.</Cell>
            </Row>
          </Table>
        </Section>
      </Document>,
    ),
  );

  // Word reads two `w:tbl` written back to back as one table and lays the
  // second out on the first one's grid. A paragraph between them is what says
  // they are two — so no two table elements may ever be adjacent.
  assertEquals(/<\/w:tbl>\s*<w:tbl[ >]/.test(xml), false);
  assertEquals((xml.match(/<w:tbl>/g) ?? []).length, 2);
});

test("a break falls through a section's suppressed heading to its first child", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a">Before.</Paragraph>
        <PageBreak id="turn" />
        <Section id="s" title="Payment" showTitle={false}>
          <Paragraph id="b">After.</Paragraph>
        </Section>
      </Document>,
    ),
  );

  assertStringIncludes(xml, '<w:pStyle w:val="PageBreakBefore"/>');
  assertEquals(xml.indexOf('w:val="PageBreakBefore"') < xml.indexOf("After."), true);
});

test("a break onto a printed heading leaves it looking like a heading", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a">Before.</Paragraph>
        <PageBreak id="turn" />
        <Section id="s" title="Payment">
          <Paragraph id="b">After.</Paragraph>
        </Section>
      </Document>,
    ),
  );

  // A paragraph names one style, so the heading cannot say both "Heading1" and
  // "carries a break". It names the style based on Heading1 instead: the page
  // turns and the heading still sets like every other section heading.
  assertStringIncludes(xml, '<w:pStyle w:val="PageBreakBeforeHeading1"/>');
  assertEquals(breaksIn(xml), 1);
  assertEquals(xml.includes('w:type="page"'), false);
  assertEquals(xml.indexOf('w:val="PageBreakBeforeHeading1"') < xml.indexOf("Payment"), true);
});

test("a break onto a printed heading keeps the heading looking like one", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D">
        <Paragraph id="a">Before.</Paragraph>
        <PageBreak id="turn" />
        <Section id="s" title="Payment">
          <Paragraph id="b">After.</Paragraph>
        </Section>
      </Document>,
    ),
  );

  // A paragraph names one style, so the heading cannot say both "Heading1" and
  // "carries a break". It names the style that is based on Heading1 instead —
  // the break happens and the heading still sets like every other heading.
  assertStringIncludes(xml, '<w:pStyle w:val="PageBreakBeforeHeading1"/>');
  assertEquals(xml.includes('w:type="page"'), false);
  assertEquals(xml.indexOf('w:val="PageBreakBeforeHeading1"') < xml.indexOf("Payment"), true);
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

test("a document says how far its running strips stand from the paper", async () => {
  const xml = await documentXml(
    await build(
      <Document
        id="d"
        title="D"
        header={<Paragraph id="head">Letterhead.</Paragraph>}
        footer={<Paragraph id="foot">Registered in England.</Paragraph>}
      >
        <Paragraph id="a">Body.</Paragraph>
      </Document>,
      {
        ...styleWithBlocks,
        page: { ...styleWithBlocks.page, headerMm: 15, footerMm: 18 },
      },
    ),
  );

  // Measured from the sheet's edge, not from the margin: 15mm is 850 twips and
  // 18mm is 1020. It is what decides whether a letterhead clears a printer's
  // unprintable edge, and it is a different distance from the margin.
  assertStringIncludes(xml, 'w:header="850"');
  assertStringIncludes(xml, 'w:footer="1020"');
});

test("a document that says nothing about them still gets Word's own distance", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D" header={<Paragraph id="head">Letterhead.</Paragraph>}>
        <Paragraph id="a">Body.</Paragraph>
      </Document>,
    ),
  );

  // 12.5mm, which is what Word's default template uses. Stated rather than left
  // to the packing library, which chose 708 twips for every document whatever
  // its margins were — a number nobody here had decided and none could change.
  assertStringIncludes(xml, 'w:header="709"');
  assertStringIncludes(xml, 'w:footer="709"');
});

test("a section prints its heading, unless it says its content already does", async () => {
  const printed = await documentXml(
    await build(
      <Document id="d" title="D">
        <Section id="s" title="Charges">
          <Paragraph id="a">A line.</Paragraph>
        </Section>
      </Document>,
    ),
  );

  assertStringIncludes(printed, 'w:val="Heading1"');
  assertStringIncludes(printed, "Charges");

  const suppressed = await build(
    <Document id="d" title="D">
      <Section id="s" title="Charges" showTitle={false}>
        <Paragraph id="a">A line.</Paragraph>
      </Section>
    </Document>,
  );
  const xml = await documentXml(suppressed);

  // Nothing printed — no heading style, no title text.
  assertEquals(xml.includes('w:val="Heading1"'), false);
  assertEquals(xml.includes("Charges"), false);
  assertStringIncludes(xml, "A line.");
  // But the section keeps its name: the model still says what it is, so ids,
  // the TOC and an engine's addressing are unchanged by not printing it.
  const section = suppressed.nodes[0];
  assertEquals(section.kind, "section");
  assertEquals(section.title, "Charges");
  assertEquals(section.kind === "section" && section.showTitle, false);
});

test("first-page furniture packs as Word's title page, not a conditional hack", async () => {
  const doc = await build(
    <Document
      id="d"
      title="D"
      header={<Paragraph id="head">INV-2026-0142</Paragraph>}
      firstHeader={false}
      footer={<Paragraph id="foot">Running foot.</Paragraph>}
      firstFooter={<Paragraph id="firstfoot">Page one foot.</Paragraph>}
    >
      <Paragraph id="a">One.</Paragraph>
      <PageBreak id="turn" />
      <Paragraph id="b">Two.</Paragraph>
    </Document>,
  );

  // The section says the first page differs.
  assertStringIncludes(await documentXml(doc), "w:titlePg");

  // Two header parts: the running one carries the strip, the first-page one
  // is present and empty — that emptiness is the whole statement.
  const names = await partNames(doc);
  const headerParts = names.filter((name) => /^word\/header\d+\.xml$/.test(name));
  assertEquals(headerParts.length, 2);
  const headerTexts = await Promise.all(headerParts.map(async (part) => {
    const xml = await partXml(doc, part);
    return [...xml.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join("");
  }));
  assertEquals(headerTexts.filter((text) => text.includes("INV-2026-0142")).length, 1);
  assertEquals(headerTexts.filter((text) => text.trim() === "").length, 1);

  // Two footer parts, each carrying its own text.
  const footerParts = names.filter((name) => /^word\/footer\d+\.xml$/.test(name));
  assertEquals(footerParts.length, 2);
  const footerTexts = await Promise.all(footerParts.map((part) => partXml(doc, part)));
  assertEquals(footerTexts.filter((xml) => xml.includes("Running foot.")).length, 1);
  assertEquals(footerTexts.filter((xml) => xml.includes("Page one foot.")).length, 1);

  // The model records the decision: `false` became an empty array, and the
  // first footer's nodes are ordinary nodes.
  assertEquals(doc.firstHeader, []);
  assertEquals(doc.firstFooter?.length, 1);
});

test("a document with no first-page furniture packs no title page", async () => {
  const xml = await documentXml(
    await build(
      <Document id="d" title="D" header={<Paragraph id="head">Strip.</Paragraph>}>
        <Paragraph id="a">One.</Paragraph>
      </Document>,
    ),
  );

  assertEquals(xml.includes("w:titlePg"), false);
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
