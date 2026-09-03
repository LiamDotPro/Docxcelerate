import { Cell, Document, Row, Table, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle, COLUMN_MM } from "../_support/style.ts";

/**
 * How a cell's words are set across it: the column's opinion, and a cell's
 * departure from it.
 *
 * Alignment is declared on the column because every row shares it — a money
 * column is right-ranged once, not once per row — and a cell may override it
 * where one row genuinely differs, which is what a totals line is.
 *
 * **One column, on purpose.** Word autofits a multi-column grid around its
 * contents (see `tables/column-widths`), so on a three-column table every
 * measured x would be the width gap rather than the alignment. A table of one
 * column has nothing to redistribute, so what the parity tier measures here is
 * alignment and only alignment.
 *
 * The cells are read the same way a paragraph is: the box is the full width of
 * the column whichever way its text is set, so only the drawn line moves.
 */
export default defineCase({
  id: "tables/cell-align",
  feature: "table.cellAlign",
  title: "A column ranged right, and a cell that departs from it",
  word: "Table → Alignment (w:jc on the cell's paragraph)",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="cell-align" title="Cell alignment">
      <Table id="figures" columns={[{ width: "auto", align: "right" }]}>
        <Row id="head" header>
          <Cell id="h1">Amount</Cell>
        </Row>
        <Row id="a">
          <Cell id="a1">Ranged right by the column</Cell>
        </Row>
        <Row id="b">
          <Cell id="b1" align="left">Ranged left by the cell</Cell>
        </Row>
        <Row id="c">
          <Cell id="c1" align="center">Centred by the cell</Cell>
        </Row>
      </Table>
    </Document>
  ),

  regions: [
    { id: "right", anchor: "Ranged right by the column" },
    { id: "left", anchor: "Ranged left by the cell" },
    { id: "centre", anchor: "Centred by the cell" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.equal(a.cell("Ranged right by the column").paragraphs[0]?.jc, "right", "the column's alignment reaches its cell's paragraph");
      is.equal(a.cell("Ranged left by the cell").paragraphs[0]?.jc, "left", "a cell that says left is written left");
      is.equal(a.cell("Centred by the cell").paragraphs[0]?.jc, "center", "and one that says centre, centred");
      is.equal(a.table(0).gridTwips.length, 1, "on a table of one column, so nothing else can move");
    },

    preview: (b, is) => {
      is.equal(b.cell("Ranged right by the column").textAlign, "right", "the preview ranges the column right");
      is.equal(b.cell("Ranged left by the cell").textAlign, "left", "and the overriding cell left");
      is.equal(b.cell("Centred by the cell").textAlign, "center", "and the centred one centre");

      // Where the words actually are, which is the half a computed style
      // cannot prove: `text-align: right` on a box that is not where you think
      // it is draws the words somewhere else.
      const cell = b.cell("Ranged right by the column");
      is.within(
        b.cellTextRight("Ranged right by the column"),
        cell.x + cell.w - cell.paddingRight,
        "0.5mm",
        "the right-ranged words end at the cell's inner right edge",
      );
      is.within(
        b.cellTextCentre("Centred by the cell"),
        cell.x + cell.w / 2,
        "0.5mm",
        "and the centred ones sit on the cell's own axis",
      );
    },

    word: (c, is) => {
      is.equal(c.cell("Ranged right by the column").alignment, "right", "Word reads the column's cells as right-ranged");
      is.equal(c.cell("Ranged left by the cell").alignment, "left", "the overriding cell as left");
      is.equal(c.cell("Centred by the cell").alignment, "center", "and the centred one as centred");
      is.within(c.cell("Ranged right by the column").width, c.mm(COLUMN_MM), "1mm", "across the whole text column");
    },

    /**
     * Where the words start, on screen against in Word.
     *
     * This is the assertion the case is for. Three alignments on one column
     * put the first glyph in three different places, and a preview that agreed
     * about which paragraph was centred while drawing it somewhere else would
     * pass every tier above.
     */
    parity: (p, is) => {
      is.within(p.previewCellX("Ranged right by the column"), p.wordCellX("Ranged right by the column"), "1mm", "the right-ranged words start where Word starts them");
      is.within(p.previewCellX("Ranged left by the cell"), p.wordCellX("Ranged left by the cell"), "1mm", "and the left-ranged ones");
      is.within(p.previewCellX("Centred by the cell"), p.wordCellX("Centred by the cell"), "1mm", "and the centred ones");
      // The first body row, not the last. Leading drifts between the two
      // engines by a fraction of a line per line — `text/line-height` is the
      // case that measures it — and by the fourth row that fraction is a
      // millimetre. Asserting it here would be this case failing for another
      // case's reason.
      is.within(p.previewCellY("Ranged right by the column"), p.wordCellY("Ranged right by the column"), "1mm", "on the line Word puts them on");
    },
  },
});
