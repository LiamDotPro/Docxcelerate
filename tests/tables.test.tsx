import { test } from "node:test";
import { assertEquals, assertStringIncludes } from "./assert.ts";
import { buildDocument, cleanMinimalDocumentStyle } from "docxcelerate";
import { createDocxBlob, createDocxDocument } from "docxcelerate/docx";
import { readPackedParagraphs } from "docxcelerate/preview";
import { documentXml, paragraphOf } from "./docx.ts";
import type {
  DocumentModel,
  DocumentNode,
  DocumentStyle,
  TableNode,
  TableRowNode,
} from "docxcelerate";
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

  const xml = await documentXml(doc as DocumentModel);

  assertStringIncludes(xml, "Subtotal");
  assertEquals(xml.indexOf("Subtotal") < xml.indexOf("Total due"), true);
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

  const rows = (await documentXml(doc as DocumentModel)).split("<w:tr>");

  // `tblHeader` is what makes Word repeat a row at the top of every page the
  // table runs onto, so it belongs to the row the table opens with and to no
  // other — a totals row carrying it would print the total on every page,
  // above the figures it adds up.
  assertStringIncludes(rows[1], "<w:tblHeader/>");
  assertStringIncludes(rows[2], '<w:tblHeader w:val="false"/>');
});

// ---------------------------------------------------------------------------
// What comes out the other end
// ---------------------------------------------------------------------------

test("a declared width comes out that wide on the page", async () => {
  const doc = await build(() => (
    <Table id="lines" columns={[{ width: "auto" }, { width: 26, align: "right" }]}>
      <Row>
        <Cell>API build</Cell>
        <Cell>6,080.00</Cell>
      </Row>
    </Table>
  ));

  const xml = await documentXml(doc as DocumentModel);

  // 26mm in twips, exactly — a money column lining up is the whole reason a
  // column was given a width rather than left to share out what is spare.
  assertStringIncludes(xml, '<w:gridCol w:w="1474"/>');
  assertStringIncludes(xml, '<w:jc w:val="right"/>');
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

// ---------------------------------------------------------------------------
// A paragraph keeps its own layout inside a cell
// ---------------------------------------------------------------------------

const inset = {
  indentMm: 10,
  indentRightMm: 10,
  border: "2F5FBD",
  borderSides: ["left"],
  borderWidthPt: 2,
  paddingPt: 6,
  fill: "EEF2FF",
} as const;

const cellStyle: DocumentStyle = {
  ...cleanMinimalDocumentStyle,
  blocks: {
    ...cleanMinimalDocumentStyle.blocks,
    signLine: { tabStopsMm: [{ at: 60, align: "right", leader: "dot" }] },
    inset: { ...inset, borderSides: ["left"] },
    standfirst: { align: "center" },
    bleeding: { bleed: true },
    measured: { maxWidthMm: 40 },
    badge: { fill: "1F2933", border: "2F5FBD" },
  },
};

const signature = () => (
  <Table id="sign" columns={[{ width: 80 }, { width: 80 }]}>
    <Row>
      <Cell id="supplier">
        <Paragraph id="s-line" variant="signLine">{"Signed\t"}</Paragraph>
        <Paragraph id="s-date" align="right">19 November 2026</Paragraph>
      </Cell>
      <Cell id="client">
        <Paragraph id="c-quote" variant="inset">Accepted on the terms above.</Paragraph>
      </Cell>
    </Row>
  </Table>
);

async function cellXml(body: () => unknown): Promise<string> {
  return documentXml({ ...(await build(body)), style: cellStyle });
}

test("a paragraph's own align wins over its column", async () => {
  const xml = await cellXml(() => (
    <Table id="t" columns={[{ width: 80 }]}>
      <Row>
        <Cell>
          <Paragraph align="right">Ranged right</Paragraph>
        </Cell>
      </Row>
    </Table>
  ));

  assertStringIncludes(paragraphOf(xml, "Ranged right"), '<w:jc w:val="right"/>');
});

test("a variant's align beats the column, and the node's beats the variant", async () => {
  const xml = await cellXml(() => (
    <Table id="t" columns={[{ width: 80, align: "right" }]}>
      <Row>
        <Cell>
          <Paragraph variant="standfirst">Centred</Paragraph>
          <Paragraph variant="standfirst" align="left">Left again</Paragraph>
        </Cell>
      </Row>
    </Table>
  ));

  assertStringIncludes(paragraphOf(xml, "Centred"), '<w:jc w:val="center"/>');
  assertStringIncludes(paragraphOf(xml, "Left again"), '<w:jc w:val="left"/>');
});

test("a plain paragraph in a cell is set by the column alone", async () => {
  const xml = await cellXml(() => (
    <Table id="t" columns={[{ width: 80, align: "right" }]}>
      <Row>
        <Cell>Plain</Cell>
      </Row>
    </Table>
  ));
  const found = paragraphOf(xml, "Plain");

  assertStringIncludes(found, '<w:jc w:val="right"/>');
  for (const absent of ["<w:ind", "<w:tabs>", "<w:pBdr>", "<w:shd"]) {
    assertEquals(found.includes(absent), false);
  }
});

test("a variant's tab stops and a tab in the text reach a cell", async () => {
  const found = paragraphOf(await cellXml(signature), "Signed");

  assertStringIncludes(found, '<w:tab w:val="right" w:pos="3402" w:leader="dot"/>');
  assertStringIncludes(found, "<w:tab/>");
});

test("a variant's indents, border and fill reach a cell paragraph", async () => {
  const found = paragraphOf(await cellXml(signature), "Accepted");

  assertStringIncludes(found, '<w:ind w:left="567" w:right="567"/>');
  assertStringIncludes(found, "<w:pBdr>");
  assertStringIncludes(found, 'w:color="2F5FBD"');
  assertStringIncludes(found, 'w:sz="16"');
  assertStringIncludes(found, 'w:space="6"');
  assertStringIncludes(found, 'w:fill="EEF2FF"');
});

test("a cell's own block stays on the cell and does not shade its paragraphs", async () => {
  const xml = await cellXml(() => (
    <Table id="t" columns={[{ width: 80 }]}>
      <Row>
        <Cell variant="badge">
          <Paragraph>Inside the badge</Paragraph>
        </Cell>
      </Row>
    </Table>
  ));
  const found = paragraphOf(xml, "Inside the badge");

  assertEquals(found.includes("<w:shd"), false);
  assertEquals(found.includes("<w:pBdr>"), false);
});

test("bleed and a measure are page measures and do nothing in a cell", async () => {
  const xml = await cellXml(() => (
    <Table id="t" columns={[{ width: 80 }]}>
      <Row>
        <Cell>
          <Paragraph variant="bleeding">Bleeds</Paragraph>
          <Paragraph variant="measured">Measured</Paragraph>
        </Cell>
      </Row>
    </Table>
  ));

  assertEquals(paragraphOf(xml, "Bleeds").includes("<w:ind"), false);
  assertEquals(paragraphOf(xml, "Measured").includes("<w:ind"), false);
});

test("a list item in a cell keeps the indent its list gives it", async () => {
  const xml = await cellXml(() => (
    <Table id="t" columns={[{ width: 80 }]}>
      <Row>
        <Cell>
          <Paragraph list="bullet">An item</Paragraph>
        </Cell>
      </Row>
    </Table>
  ));

  assertStringIncludes(paragraphOf(xml, "An item"), '<w:ind w:left="360" w:hanging="360"/>');
});

test("the preview reader finds a cell's tab stops, and the body after the table", async () => {
  const doc = await build(() => (
    <>
      {signature()}
      <Paragraph id="after">After the table</Paragraph>
    </>
  ));
  const blob = await createDocxBlob({ ...doc, style: cellStyle });
  const read = await readPackedParagraphs(new Uint8Array(await blob.arrayBuffer()));
  const signed = read.find((entry) => entry.text.includes("Signed"));
  const stop = signed?.tabStops[0];

  assertEquals(stop?.align, "right");
  assertEquals(stop?.leader, "dot");
  assertEquals(Math.abs((stop?.positionPt ?? 0) - 60 * 72 / 25.4) < 0.1, true);
  assertEquals(read.at(-1)?.text.includes("After the table"), true);
});
