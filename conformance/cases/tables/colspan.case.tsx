import { Cell, Document, Row, Table, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle, COLUMN_MM } from "../_support/style.ts";

/**
 * A cell that runs across more than one column.
 *
 * The shape a totals line is: the figures keep their columns and the label
 * above them takes the rest of the row. In OOXML that is `w:gridSpan`, and the
 * thing worth checking is not that the attribute was written but that the
 * *grid* still adds up — a row of two cells where one spans two of three
 * columns has to occupy the same three columns as the row above it, or Word
 * quietly rebuilds the table around the mismatch.
 *
 * So the assertions are about columns rather than about siblings. `column` is
 * where a cell begins in the grid, counting spans: on the spanning row the
 * second cell begins at column 2, not at column 1, and a probe that counted
 * siblings would report the same number for both and prove nothing.
 *
 * The spanning cell covers the whole table, which is deliberate. Word autofits
 * a grid it was not told to keep fixed (`tables/column-widths`), so a cell
 * spanning *some* of the columns would be measured against widths Word had
 * already changed. One that spans all of them is the width of the table in
 * both engines whatever autofit does with the boundaries inside it.
 */
export default defineCase({
  id: "tables/colspan",
  feature: "table.colspan",
  title: "A cell that runs across more than one column",
  word: "Table → Merge Cells, horizontally (w:gridSpan)",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="colspan" title="Horizontal span">
      <Table id="charges" columns={[{ width: "auto" }, { width: 30, align: "right" }]}>
        <Row id="head" header>
          <Cell id="h1">Description</Cell>
          <Cell id="h2">Amount</Cell>
        </Row>
        <Row id="line">
          <Cell id="c1">Consultancy</Cell>
          <Cell id="c2">1,250.00</Cell>
        </Row>
        <Row id="note">
          <Cell id="n1" span={2}>Spanning both columns, as a note under the figures does</Cell>
        </Row>
      </Table>
    </Document>
  ),

  regions: [
    { id: "label", anchor: "Consultancy" },
    { id: "money", anchor: "1,250.00" },
    { id: "span", anchor: "Spanning both columns" },
  ],

  expect: {
    ooxml: (a, is) => {
      const table = a.table(0);

      is.equal(table.gridTwips.length, 2, "the grid is two columns wide");
      is.equal(table.row(1).cellCount, 2, "an ordinary row holds two cells");
      is.equal(table.row(2).cellCount, 1, "the spanning row holds one");

      is.equal(a.cell("Spanning both columns").gridSpan, 2, "and that one declares w:gridSpan 2");
      is.equal(a.cell("Consultancy").gridSpan, 1, "where a plain cell declares none");
      is.equal(a.cell("1,250.00").column, 1, "the money cell begins in the second column");
      is.equal(a.cell("Spanning both columns").column, 0, "and the spanning one in the first");
    },

    preview: (b, is) => {
      is.equal(b.cell("Spanning both columns").colSpan, 2, "the preview draws the cell across two columns");
      is.equal(b.cell("Consultancy").colSpan, 1, "and a plain one across one");
      is.within(
        b.cell("Spanning both columns").w,
        b.mm(COLUMN_MM),
        "1mm",
        "so it is as wide as the whole table",
      );
      is.within(
        b.cell("Spanning both columns").w,
        b.cell("Consultancy").w + b.cell("1,250.00").w,
        "0.5mm",
        "which is the two columns it covers, added up",
      );
    },

    word: (c, is) => {
      // Uniform is Word's word for "every row has the same cells in the same
      // places". A spanned row is exactly what makes a table not uniform, so
      // this is the fact rather than a symptom of one.
      is.equal(c.table(0).uniform, false, "Word reads the table as no longer a plain grid");
      is.within(
        c.cell("Spanning both columns").width,
        c.mm(COLUMN_MM),
        "1mm",
        "with the spanning cell the width of the whole table",
      );
      is.equal(c.cell("Spanning both columns").column, 0, "beginning in the first column");
      is.equal(c.cell("1,250.00").column, 1, "and the money cell still in the second");
    },

    parity: (p, is) => {
      is.within(p.previewCellX("Spanning both columns"), p.wordCellX("Spanning both columns"), "1mm", "the spanning cell's words start where Word starts them");
      is.within(p.previewCellWidth("Spanning both columns"), p.wordCellWidth("Spanning both columns"), "1mm", "and it is as wide on screen as in Word");
    },
  },
});
