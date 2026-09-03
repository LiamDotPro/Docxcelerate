import { Cell, Document, Paragraph, Row, Table, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * A table inside a cell of another table.
 *
 * Listed as untested rather than as supported or absent, which is a different
 * kind of entry: nothing in the renderer forbids it — a cell renders whatever
 * nodes it holds and a table is a node — but nothing had ever asked it to, and
 * "it should work" is not a measurement. This case is what turns that row of
 * the inventory into a fact.
 *
 * There is one thing here worth more than the answer, and it is why the probes
 * had to change before this case could exist. A table is the only shape in a
 * document that contains itself, so every reading of one is a chance to read
 * the wrong one. `element(cellXml, "w:tcPr")` returns the *inner* cell's
 * properties for an outer cell that holds a nested table, and a non-greedy
 * `<w:tbl>…</w:tbl>` ends halfway through the outer table. Both mistakes give
 * a plausible number for the wrong element. The assertions below are as much
 * about the probes reading the right table as about the packer writing one.
 *
 * The empty paragraph after the inner table is not a stray. OOXML requires a
 * cell to *end* in a `w:p`, so a cell whose last child is a table gets one
 * written after it — which is why the outer cell reports two paragraphs for
 * the one it was given.
 */
export default defineCase({
  id: "tables/nested",
  feature: "table.nested",
  title: "A table inside a cell of another table",
  word: "A w:tbl inside a w:tc",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="nested" title="Nested table">
      <Table id="outer" columns={[{ width: 60 }, { width: "auto" }]}>
        <Row id="head" header>
          <Cell id="h1">Party</Cell>
          <Cell id="h2">Terms</Cell>
        </Row>
        <Row id="body">
          <Cell id="c1">
            <Paragraph id="lead">A cell holding a table of its own:</Paragraph>
            <Table id="inner" columns={[{ width: 30 }, { width: 30, align: "right" }]}>
              <Row id="i1">
                <Cell id="i1c1">Inner left</Cell>
                <Cell id="i1c2">Inner right</Cell>
              </Row>
              <Row id="i2">
                <Cell id="i2c1">Second inner</Cell>
                <Cell id="i2c2">Second right</Cell>
              </Row>
            </Table>
          </Cell>
          <Cell id="c2">The outer cell beside it</Cell>
        </Row>
      </Table>
    </Document>
  ),

  regions: [
    { id: "outer", anchor: "The outer cell" },
    { id: "inner", anchor: "Inner left" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.equal(a.tableCount, 1, "the body holds one table");
      is.equal(a.allTables().length, 2, "which holds a second inside it");
      // Innermost first, which is the order an anchor is resolved in.
      is.equal(a.allTables()[0].path, "0.1.0.0", "in the first cell of the outer table's second row");
      is.equal(a.allTables()[0].depth, 1, "one table deep");
      is.equal(a.allTables()[1].path, "0", "with the table that holds it last");

      const outer = a.table(0);
      is.equal(outer.rowCount, 2, "the outer table has two rows");
      is.equal(outer.row(1).cellCount, 2, "and two cells across the second");

      // The reading that a first-match probe gets wrong. The outer cell holds
      // a nested table whose cells have properties of their own; these are the
      // outer cell's.
      const cell = outer.row(1).cell(0);
      is.equal(cell.tables.length, 1, "the second cell holds one table");
      is.equal(cell.paragraphs.length, 2, "and two paragraphs of its own — the one it was given, and the empty one a cell has to end in");
      is.equal(cell.paragraphs[1].text, "", "which is the empty one");
      is.equal(cell.tables[0].rowCount, 2, "the inner table has two rows");
      is.equal(cell.tables[0].gridTwips.length, 2, "on a grid of two columns");

      is.equal(a.cell("Inner right").paragraphs[0]?.jc, "right", "and the inner table's own column alignment survives");
    },

    preview: (b, is) => {
      is.equal(b.tables.length, 2, "the preview draws both tables");
      is.equal(b.nestedTables().length, 1, "one of them inside the other");
      is.equal(b.table(1).nested, true, "which is the second one it met");

      // The inner table is drawn inside the outer cell's box, which is the
      // whole claim: a nested table that escaped its cell would still be two
      // tables and would still be nested by every reading above.
      const outer = b.cell("A cell holding a table");
      const inner = b.table(1);
      is.greater(inner.x, outer.x - 1, "and drawn inside its cell's left edge");
      is.greater(outer.x + outer.w + 1, inner.x + inner.w, "and inside its right");
    },

    word: (c, is) => {
      is.equal(c.tables.length, 1, "Word finds one table in the body");
      is.equal(c.table(0).nested.length, 1, "with one nested inside it");
      is.equal(c.table(0).nested[0].rowCount, 2, "and two rows in that one");
      is.equal(c.cell("Inner right").alignment, "right", "reading the inner column as right-ranged");
    },

    parity: (p, is) => {
      is.within(p.previewCellX("Inner left"), p.wordCellX("Inner left"), "1mm", "the inner table's words start where Word starts them");
      is.within(p.previewCellY("Inner left"), p.wordCellY("Inner left"), "1mm", "on the line Word puts them on");
    },
  },
});
