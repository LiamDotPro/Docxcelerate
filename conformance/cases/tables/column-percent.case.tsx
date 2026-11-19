import { Cell, Document, Row, Table, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle, COLUMN_MM, TWIPS_PER_MM } from "../_support/style.ts";

/**
 * Columns declared as a share of the text column.
 *
 * A percentage is worked out into fixed twips when the page is known, not
 * written as `w:type="pct"`: `tables/column-widths` documents why the grid has
 * to be fixed twips for Word to honour it. So the claim is the same as there,
 * reached from a proportion instead of a measurement.
 */
const SHARES = [0.5, 0.3, 0.2];

export default defineCase({
  id: "tables/column-percent",
  feature: "table.columnPercent",
  title: "Columns declared as a share of the text column",
  word: "Table Properties → Column → Preferred width (w:tblGrid, w:tblLayout)",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="column-percent" title="Column percent">
      <Table id="charges" columns={[{ width: "50%" }, { width: "30%" }, { width: "20%" }]}>
        <Row id="line">
          <Cell id="c1">Consultancy</Cell>
          <Cell id="c2">1,250.00</Cell>
          <Cell id="c3">Fixed fee</Cell>
        </Row>
      </Table>
    </Document>
  ),

  regions: [
    { id: "first", anchor: "Consultancy" },
    { id: "money", anchor: "1,250.00" },
    { id: "last", anchor: "Fixed fee" },
  ],

  expect: {
    ooxml: (a, is) => {
      const table = a.table(0);

      is.equal(table.gridTwips.length, 3, "the grid declares three columns");
      SHARES.forEach((share, i) =>
        is.within(table.gridTwips[i], COLUMN_MM * TWIPS_PER_MM * share, 1, `column ${i + 1} is ${share * 100}% of the column`)
      );
      is.equal(table.width?.type, "dxa", "in twips, not as a percentage of something");
      is.equal(table.layout, "fixed", "and declares a fixed layout, so the grid is not a suggestion");
    },

    preview: (b, is) => {
      is.within(b.cell("Consultancy").w, b.mm(COLUMN_MM * SHARES[0]), "1mm", "the first column draws half the column");
      is.within(b.cell("1,250.00").w, b.mm(COLUMN_MM * SHARES[1]), "1mm", "the second draws 30%");
      is.within(b.cell("Fixed fee").w, b.mm(COLUMN_MM * SHARES[2]), "1mm", "the third draws 20%");
    },

    word: (c, is) => {
      is.within(c.cell("Consultancy").width, c.mm(COLUMN_MM * SHARES[0]), "1mm", "Word makes the first half the column");
      is.within(c.cell("1,250.00").width, c.mm(COLUMN_MM * SHARES[1]), "1mm", "the second 30%");
      is.within(c.cell("Fixed fee").width, c.mm(COLUMN_MM * SHARES[2]), "1mm", "and the third 20%");
    },

    parity: (p, is) => {
      is.within(p.previewCellX("1,250.00"), p.wordCellX("1,250.00"), "1mm", "the second column starts where Word starts it");
      is.within(p.previewCellX("Fixed fee"), p.wordCellX("Fixed fee"), "1mm", "and so does the third");
      is.within(p.previewCellWidth("1,250.00"), p.wordCellWidth("1,250.00"), "1mm", "the second is as wide on screen as in Word");
    },
  },
});
