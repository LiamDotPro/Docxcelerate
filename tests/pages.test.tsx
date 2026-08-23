import { test } from "node:test";
import { assertEquals, assertStringIncludes } from "./assert.ts";
import { buildDocument } from "docxcelerate";
import { createDocxDocument } from "docxcelerate/docx";
import { renderDocumentWebsite } from "docxcelerate/renderer";
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
 * is what these check.
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

function build(node: unknown, style?: DocumentStyle) {
  return buildDocument(
    template<Record<string, never>>(node as never),
    {},
    { branchMode: "decide", dynamicMode: "placeholder" },
  ).then((doc) => (style ? { ...doc, style } : doc));
}

function pagesIn(html: string): number {
  return (html.match(/class="a4-page"/g) ?? []).length;
}

// ---------------------------------------------------------------------------
// Where a page ends
// ---------------------------------------------------------------------------

test("a document with no break is one page", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Paragraph id="a">One.</Paragraph>
      <Paragraph id="b">Two.</Paragraph>
    </Document>,
  );

  assertEquals(pagesIn(renderDocumentWebsite(doc)), 1);
});

test("a break starts the next page, and is not printed itself", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Paragraph id="owed">What is owed.</Paragraph>
      <PageBreak id="turn" />
      <Paragraph id="pay">How to pay it.</Paragraph>
    </Document>,
  );

  const html = renderDocumentWebsite(doc);
  const [first, second] = html.split('class="a4-page"').slice(1);

  assertEquals(pagesIn(html), 2);
  assertStringIncludes(first, "What is owed.");
  assertEquals(first.includes("How to pay it."), false);
  assertStringIncludes(second, "How to pay it.");
});

test("a trailing break does not leave a sheet with nothing on it", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Paragraph id="a">One.</Paragraph>
      <PageBreak id="turn" />
    </Document>,
  );

  assertEquals(pagesIn(renderDocumentWebsite(doc)), 1);
});

// ---------------------------------------------------------------------------
// What repeats
// ---------------------------------------------------------------------------

test("the running header and footer are drawn on every page", async () => {
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

  const html = renderDocumentWebsite(doc);

  assertEquals((html.match(/class="page-header"/g) ?? []).length, 2);
  assertEquals((html.match(/class="page-footer"/g) ?? []).length, 2);
  assertEquals((html.match(/INV-2026-0142/g) ?? []).length, 2);
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

test("the title is the document's name, so it appears once and not per page", async () => {
  const doc = await build(
    <Document id="d" title="Invoice">
      <Paragraph id="a">One.</Paragraph>
      <PageBreak id="turn" />
      <Paragraph id="b">Two.</Paragraph>
    </Document>,
  );

  const html = renderDocumentWebsite(doc);

  assertEquals((html.match(/class="document-title"/g) ?? []).length, 1);
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

  const html = renderDocumentWebsite(doc);

  assertEquals(html.includes('<h1 class="document-title"'), false);
  assertEquals(doc.title, "Invoice");
  assertEquals(typeof createDocxDocument(doc as DocumentModel), "object");
});

test("a document that says nothing about its title still gets one", async () => {
  const doc = await build(
    <Document id="d" title="Invoice">
      <Paragraph id="a">One.</Paragraph>
    </Document>,
    styleWithBlocks,
  );

  assertStringIncludes(renderDocumentWebsite(doc), '<h1 class="document-title"');
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
  assertStringIncludes(renderDocumentWebsite(doc), 'data-separator=" of "');
});

test("each page knows its own number, which is what the count is drawn from", async () => {
  const doc = await build(
    <Document id="d" title="D" footer={<PageNumber id="n" />}>
      <Paragraph id="a">One.</Paragraph>
      <PageBreak id="turn" />
      <Paragraph id="b">Two.</Paragraph>
    </Document>,
  );

  const html = renderDocumentWebsite(doc);

  assertStringIncludes(html, "--page-current:'1'");
  assertStringIncludes(html, "--page-current:'2'");
  assertStringIncludes(html, "--page-total:'2'");
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
  const doc = await build(
    <Document id="d" title="D">
      <Paragraph id="a" variant="band">Dates.</Paragraph>
    </Document>,
    styleWithBlocks,
  );

  const html = renderDocumentWebsite(doc);

  assertStringIncludes(html, '[data-variant="band"]');
  assertStringIncludes(html, "background: #F4F6FD;");
  assertStringIncludes(html, 'data-variant="band"');
});

test("a variant the theme has never heard of draws as an ordinary block", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Paragraph id="a" variant="nonesuch">Text.</Paragraph>
    </Document>,
    styleWithBlocks,
  );

  const html = renderDocumentWebsite(doc);

  assertEquals(html.includes('[data-variant="nonesuch"]'), false);
  assertStringIncludes(html, "Text.");
});

test("a bleeding block escapes the margins, so a band crosses the whole sheet", async () => {
  const bleeding: DocumentStyle = {
    ...styleWithBlocks,
    blocks: { band: { fill: "F4F6FD", bleed: true } },
  };
  const doc = await build(
    <Document id="d" title="D">
      <Paragraph id="a" variant="band">Dates.</Paragraph>
    </Document>,
    bleeding,
  );

  assertStringIncludes(
    renderDocumentWebsite(doc),
    "margin-left: calc(-1 * var(--page-margin-left));",
  );
});

test("a block that does not bleed stays inside the text column", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Paragraph id="a" variant="band">Dates.</Paragraph>
    </Document>,
    styleWithBlocks,
  );

  assertEquals(renderDocumentWebsite(doc).includes("calc(-1 * var(--page-margin-left))"), false);
});

test("a filled cell drops the rule that separates rows, so a band is one strip", async () => {
  const doc = await build(
    <Document id="d" title="D">
      <Table id="t" columns={[{ width: "auto" }, { width: 30 }]}>
        <Row>
          <Cell variant="band">Issue date</Cell>
          <Cell variant="band">Due date</Cell>
        </Row>
      </Table>
    </Document>,
    styleWithBlocks,
  );

  assertStringIncludes(renderDocumentWebsite(doc), 'td[data-variant="band"]');
});

// ---------------------------------------------------------------------------
// What Word is handed
// ---------------------------------------------------------------------------

test("the DOCX packer takes a break, a header, a footer and a page number", async () => {
  const doc = await build(
    <Document
      id="d"
      title="D"
      header={<Paragraph id="head">INV-2026-0142</Paragraph>}
      footer={<PageNumber id="n" />}
    >
      <Paragraph id="a" variant="band">One.</Paragraph>
      <PageBreak id="turn" />
      <Paragraph id="b">Two.</Paragraph>
    </Document>,
    styleWithBlocks,
  );

  // Packing is the assertion: docx refuses a malformed header, footer or field.
  assertEquals(typeof createDocxDocument(doc as DocumentModel), "object");
});
