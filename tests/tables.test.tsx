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
  TableCellNode,
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

test("the model carries an exact height", async () => {
  const doc = await build(() => (
    <Table id="t" columns={[{ width: "auto" }]}>
      <Row height={40} heightRule="exact">
        <Cell>x</Cell>
      </Row>
    </Table>
  ));
  const row = rowsOf(tableOf(doc as DocumentModel))[0];

  assertEquals(row.height, 40);
  assertEquals(row.heightRule, "exact");
});

test("the default rule is not carried", async () => {
  const doc = await build(() => (
    <Table id="t" columns={[{ width: "auto" }]}>
      <Row height={40}>
        <Cell>x</Cell>
      </Row>
    </Table>
  ));
  const row = rowsOf(tableOf(doc as DocumentModel))[0];

  assertEquals(row.height, 40);
  assertEquals("heightRule" in row, false);
});

test("a rule with no height, and a plain row, carry nothing", async () => {
  const doc = await build(() => (
    <Table id="t" columns={[{ width: "auto" }]}>
      <Row heightRule="exact">
        <Cell>x</Cell>
      </Row>
      <Row>
        <Cell>y</Cell>
      </Row>
    </Table>
  ));

  for (const row of rowsOf(tableOf(doc as DocumentModel))) {
    assertEquals("height" in row, false);
    assertEquals("heightRule" in row, false);
  }
});

test("a row's height is written in twips, with its rule", async () => {
  const doc = await build(() => (
    <Table id="t" columns={[{ width: "auto" }]}>
      <Row height={40} heightRule="exact">
        <Cell>a</Cell>
      </Row>
      <Row height={56.7}>
        <Cell>b</Cell>
      </Row>
      <Row>
        <Cell>c</Cell>
      </Row>
    </Table>
  ));
  const rows = (await documentXml(doc as DocumentModel)).split("<w:tr>");

  assertStringIncludes(rows[1], '<w:trHeight w:val="800" w:hRule="exact"/>');
  assertStringIncludes(rows[2], '<w:trHeight w:val="1134" w:hRule="atLeast"/>');
  assertEquals(rows[3].includes("w:trHeight"), false);
});

test("a header row keeps both its repeat and its height", async () => {
  const doc = await build(() => (
    <Table id="t" columns={[{ width: "auto" }]}>
      <Row header height={30}>
        <Cell>h</Cell>
      </Row>
    </Table>
  ));
  const rows = (await documentXml(doc as DocumentModel)).split("<w:tr>");

  assertStringIncludes(rows[1], "<w:tblHeader/>");
  assertStringIncludes(rows[1], '<w:trHeight w:val="600" w:hRule="atLeast"/>');
});

// ---------------------------------------------------------------------------
// Rows that stay whole
// ---------------------------------------------------------------------------

test("the model carries keepTogether only when asked", async () => {
  const doc = await build(() => (
    <Table id="t" columns={[{ width: "auto" }]}>
      <Row keepTogether>
        <Cell>Kept</Cell>
      </Row>
      <Row>
        <Cell>Free</Cell>
      </Row>
    </Table>
  ));

  assertEquals(rowsOf(tableOf(doc))[0].keepTogether, true);
  assertEquals("keepTogether" in rowsOf(tableOf(doc))[1], false);
});

test("w:cantSplit is written on the marked rows and on no other", async () => {
  const doc = await build(() => (
    <Table id="t" columns={[{ width: "auto" }]}>
      <Row header keepTogether>
        <Cell>Description</Cell>
      </Row>
      <Row keepTogether>
        <Cell>API build</Cell>
      </Row>
      <Row>
        <Cell>Support</Cell>
      </Row>
    </Table>
  ));

  const rows = (await documentXml(doc as DocumentModel)).split("<w:tr>");

  assertStringIncludes(rows[1], "<w:cantSplit/>");
  assertStringIncludes(rows[1], "<w:tblHeader/>");
  assertStringIncludes(rows[2], "<w:cantSplit/>");
  assertEquals(rows[3].includes("w:cantSplit"), false);
});

test("a row produced by a `.map()` carries keepTogether to every pass", async () => {
  const doc = await build(() => (
    <Table id="t" columns={[{ width: "auto" }, { width: 26 }]}>
      {lines.map((line) => (
        <Row keepTogether>
          <Cell>{line.desc}</Cell>
          <Cell>{line.amount}</Cell>
        </Row>
      ))}
    </Table>
  ));

  assertEquals(rowsOf(tableOf(doc)).every((row) => row.keepTogether === true), true);
  const xml = await documentXml(doc as DocumentModel);
  assertEquals(xml.split("<w:cantSplit/>").length - 1, lines.length);
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
// Where a table sits across the column
// ---------------------------------------------------------------------------

function narrow(align?: "left" | "center" | "right") {
  return build(() => (
    <Table id="t" align={align} columns={[{ width: 60 }]}>
      <Row>
        <Cell id="c">Sixty millimetres</Cell>
      </Row>
    </Table>
  ));
}

async function tblPrOf(align?: "left" | "center" | "right"): Promise<string> {
  const xml = await documentXml(await narrow(align));

  return xml.match(/<w:tblPr>[\s\S]*?<\/w:tblPr>/)?.[0] ?? "";
}

test("a table keeps the align it was given", async () => {
  assertEquals(tableOf(await narrow("right")).align, "right");
});

test("a table that does not say where it sits carries no align", async () => {
  assertEquals("align" in tableOf(await narrow()), false);
});

test("a centred table writes w:jc on the table's own properties", async () => {
  assertStringIncludes(await tblPrOf("center"), '<w:tblW w:type="dxa" w:w="3402"/><w:jc w:val="center"/>');
});

test("a table ranged right writes w:jc right", async () => {
  assertStringIncludes(await tblPrOf("right"), '<w:jc w:val="right"/>');
});

test("a table with no align writes no w:jc", async () => {
  assertEquals((await tblPrOf()).includes("<w:jc"), false);
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

const turned = () => (
  <Table id="t" columns={[{ width: 12 }, { width: 12 }, { width: "auto" }]}>
    <Row>
      <Cell textDirection="up">Up</Cell>
      <Cell textDirection="down">Down</Cell>
      <Cell>Across</Cell>
    </Row>
  </Table>
);

test("the model carries a direction only when it is set", async () => {
  const doc = await build(turned);
  const cells = rowsOf(tableOf(doc as DocumentModel))[0].children as TableCellNode[];

  assertEquals(cells[0].textDirection, "up");
  assertEquals(cells[1].textDirection, "down");
  assertEquals("textDirection" in cells[2], false);
});

test("a direction is written into the cell's properties", async () => {
  const cells = (await documentXml(await build(turned))).split("<w:tc>");

  assertStringIncludes(cells[1], '<w:textDirection w:val="btLr"/>');
  assertStringIncludes(cells[2], '<w:textDirection w:val="tbRl"/>');
  assertEquals(cells[3].includes("w:textDirection"), false);
});

test("a turned cell keeps its other properties", async () => {
  const doc = await build(() => (
    <Table id="t" columns={[{ width: 12 }, { width: 12 }]}>
      <Row header>
        <Cell textDirection="up" span={2}>Week one</Cell>
      </Row>
    </Table>
  ));
  const cell = (await documentXml(doc)).split("<w:tc>")[1];

  assertStringIncludes(cell, '<w:gridSpan w:val="2"/>');
  assertStringIncludes(cell, "<w:shd ");
  assertStringIncludes(cell, '<w:textDirection w:val="btLr"/>');
  if (cell.includes("<w:vAlign")) {
    assertEquals(cell.indexOf("w:textDirection") < cell.indexOf("<w:vAlign"), true);
  }
});
