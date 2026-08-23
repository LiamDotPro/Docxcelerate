import { test } from "node:test";
import { assertEquals, assertStringIncludes } from "./assert.ts";
import { buildDocument } from "docxcelerate";
import { createDocxDocument } from "docxcelerate/docx";
import { renderDocumentWebsite } from "docxcelerate/renderer";
import type { DocumentModel, DocumentNode, TableNode, TableRowNode } from "docxcelerate";
import { Cell, Document, Paragraph, Row, Section, Table, template } from "docxcelerate/template";

/**
 * What a table has to be for an invoice to be one.
 *
 * The columns belong to the table; everything else is a node, which is what
 * lets a loop produce rows and a condition drop one. The cases below are the
 * places this could be got wrong in a way nobody notices until a document is
 * printed: a total above the figures it adds up, a money column that does not
 * line up, a two-line cell that prints its two lines as one.
 */

interface Line {
  desc: string;
  amount: string;
}

const lines: Line[] = [
  { desc: "API build", amount: "6,080.00" },
  { desc: "Support retainer", amount: "950.00" },
];

function build(body: () => unknown) {
  const Body = body as () => never;

  return buildDocument(
    template<{ lines: Line[] }>(
      <Document id="doc" title="Doc">
        <Section id="body" title="Body">
          <Body />
        </Section>
      </Document>,
    ),
    { lines },
    { branchMode: "decide", dynamicMode: "placeholder" },
  );
}

function tableOf(doc: DocumentModel): TableNode {
  const section = doc.nodes[0];

  if (section?.kind !== "section") {
    throw new Error("expected the body section");
  }

  const table = section.children[0];

  if (table?.kind !== "table") {
    throw new Error(`expected a table, got ${table?.kind ?? "nothing"}`);
  }

  return table;
}

function rowsOf(node: TableNode): TableRowNode[] {
  return node.children.filter((child): child is TableRowNode => child.kind === "tableRow");
}

function textsOf(node: DocumentNode): string[] {
  if (node.kind === "paragraph") {
    return [node.text ?? ""];
  }

  return "children" in node ? node.children.flatMap(textsOf) : [];
}

// ---------------------------------------------------------------------------
// The shape a table carries
// ---------------------------------------------------------------------------

test("the columns are the table's, and every row shares them", async () => {
  const doc = await build(() => (
    <Table id="t" columns={[{ width: "auto" }, { width: 26, align: "right" }]}>
      <Row header>
        <Cell>Description</Cell>
        <Cell>Amount</Cell>
      </Row>
    </Table>
  ));

  assertEquals(tableOf(doc).columns, [{ width: "auto" }, { width: 26, align: "right" }]);
});

test("a cell given text becomes the paragraph it would have had to be written as", async () => {
  const doc = await build(() => (
    <Table id="t" columns={[{ width: "auto" }]}>
      <Row>
        <Cell>950.00</Cell>
      </Row>
    </Table>
  ));

  assertEquals(textsOf(tableOf(doc)), ["950.00"]);
});

test("a cell given paragraphs keeps them apart, because two lines are two lines", async () => {
  // The description above a muted note. Flattening these into one paragraph
  // would lose the line the design puts the note on.
  const doc = await build(() => (
    <Table id="t" columns={[{ width: "auto" }]}>
      <Row>
        <Cell>
          <Paragraph>API development</Paragraph>
          <Paragraph>Repairs booking endpoints</Paragraph>
        </Cell>
      </Row>
    </Table>
  ));

  assertEquals(textsOf(tableOf(doc)), ["API development", "Repairs booking endpoints"]);
});

test("a `.map()` over the data produces a row per entry", async () => {
  const doc = await build(() => (
    <Table id="t" columns={[{ width: "auto" }, { width: 26 }]}>
      <Row header>
        <Cell>Description</Cell>
        <Cell>Amount</Cell>
      </Row>
      {lines.map((line) => (
        <Row>
          <Cell>{line.desc}</Cell>
          <Cell>{line.amount}</Cell>
        </Row>
      ))}
    </Table>
  ));

  assertEquals(rowsOf(tableOf(doc)).length, 3);
  assertEquals(textsOf(tableOf(doc)).slice(2), [
    "API build",
    "6,080.00",
    "Support retainer",
    "950.00",
  ]);
});

test("rows produced by a loop are named apart, so no id is used twice", async () => {
  const doc = await build(() => (
    <Table id="t" columns={[{ width: "auto" }]}>
      {lines.map((line) => (
        <Row>
          <Cell>{line.desc}</Cell>
        </Row>
      ))}
    </Table>
  ));

  const ids = rowsOf(tableOf(doc)).map((row) => row.id);

  assertEquals(new Set(ids).size, ids.length);
});

// ---------------------------------------------------------------------------
// Where a header row belongs
// ---------------------------------------------------------------------------

test("a totals row stays under the figures it adds up", async () => {
  // A header row is drawn as one, but only the rows a table opens with are its
  // heading. Lifting a marked row from the bottom would print the total above
  // the subtotal — which is exactly what this did before it was fixed.
  const doc = await build(() => (
    <Table id="t" columns={[{ width: "auto" }, { width: 26 }]}>
      <Row>
        <Cell>Subtotal</Cell>
        <Cell>18,650.00</Cell>
      </Row>
      <Row header>
        <Cell>Total due</Cell>
        <Cell>22,380.00</Cell>
      </Row>
    </Table>
  ));

  const html = renderDocumentWebsite(doc);
  const body = html.slice(html.indexOf("<tbody>"));

  assertStringIncludes(body, "Subtotal");
  assertEquals(body.indexOf("Subtotal") < body.indexOf("Total due"), true);
});

test("the heading a table opens with is the one that repeats across pages", async () => {
  const doc = await build(() => (
    <Table id="t" columns={[{ width: "auto" }]}>
      <Row header>
        <Cell>Description</Cell>
      </Row>
      <Row>
        <Cell>API build</Cell>
      </Row>
    </Table>
  ));

  const head = renderDocumentWebsite(doc).match(/<thead>[\s\S]*?<\/thead>/)?.[0] ?? "";

  assertStringIncludes(head, "Description");
  assertEquals(head.includes("API build"), false);
});

// ---------------------------------------------------------------------------
// What comes out the other end
// ---------------------------------------------------------------------------

test("the preview draws a real table, with the declared widths on the columns", async () => {
  const doc = await build(() => (
    <Table id="lines" columns={[{ width: "auto" }, { width: 26, align: "right" }]}>
      <Row>
        <Cell>API build</Cell>
        <Cell>6,080.00</Cell>
      </Row>
    </Table>
  ));

  const html = renderDocumentWebsite(doc);

  assertStringIncludes(html, `<table class="doc-table" data-node-id="lines">`);
  assertStringIncludes(html, `<col style="width:26mm">`);
  assertStringIncludes(html, "text-align:right");
});

test("the DOCX packer accepts a table rather than falling through to text", async () => {
  const doc = await build(() => (
    <Table id="t" columns={[{ width: "auto" }, { width: 26, align: "right" }]}>
      <Row header>
        <Cell>Description</Cell>
        <Cell>Amount</Cell>
      </Row>
      {lines.map((line) => (
        <Row>
          <Cell>{line.desc}</Cell>
          <Cell>{line.amount}</Cell>
        </Row>
      ))}
    </Table>
  ));

  // Packing is the assertion: docx rejects a malformed table, and an empty
  // cell among them, so a document that packs is a table Word will open.
  assertEquals(typeof createDocxDocument(doc), "object");
});

test("a cell with nothing in it still packs, because Word will not take an empty one", async () => {
  const doc = await build(() => (
    <Table id="t" columns={[{ width: "auto" }, { width: 26 }]}>
      <Row>
        <Cell>Balance</Cell>
        <Cell></Cell>
      </Row>
    </Table>
  ));

  assertEquals(typeof createDocxDocument(doc), "object");
});
