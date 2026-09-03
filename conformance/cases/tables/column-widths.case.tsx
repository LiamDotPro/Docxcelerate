import { Cell, Document, Row, Table, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle, COLUMN_MM, TWIPS_PER_MM } from "../_support/style.ts";

/**
 * The columns, declared once and expected to come out that wide.
 *
 * This is the reason a table is a table here rather than tabbed text. A money
 * column given 40mm has to be 40mm on the page, in the preview and in Word, or
 * the figures under it do not line up — and lining up is the whole of what the
 * reader is being shown. The `"auto"` column is the other half of the claim:
 * whatever the fixed ones leave is shared out, so the table fills the text
 * column exactly and never overhangs the margin by one conversion's rounding.
 *
 * **This case was red in Word when it was written, and finding that is what it
 * was for.** The packer emitted a correct `w:tblGrid` and no `w:tblLayout` —
 * and a table without one is autofit to its contents, so Word reworked the
 * grid around the words in the cells and printed 60mm / 40mm / 70mm as 67.5mm
 * / 50.9mm / 51.6mm. The total was right, which is exactly why nobody had
 * noticed: the table still filled the column, and only the boundaries inside
 * it were wrong. The preview honoured the grid, so the two engines disagreed
 * by up to 18mm about where a column starts. `renderTable` now declares
 * `TableLayoutType.FIXED` and the assertion below is what keeps it there.
 *
 * Watch the last twip. 170mm is 9637.795 twips and the packer works in whole
 * ones, so the widths are asserted to a twip of slack rather than to the
 * number, and their sum against the table's own declared width rather than
 * against the arithmetic that produced both.
 */
const FIXED_MM = [60, 40];
const AUTO_MM = COLUMN_MM - FIXED_MM[0] - FIXED_MM[1];

export default defineCase({
  id: "tables/column-widths",
  feature: "table.columnWidths",
  title: "Columns declared in millimetres, and one that takes what is left",
  word: "Table Properties → Column → Preferred width (w:tblGrid, w:tblLayout)",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="column-widths" title="Column widths">
      <Table
        id="charges"
        columns={[{ width: FIXED_MM[0] }, { width: FIXED_MM[1] }, { width: "auto" }]}
      >
        <Row id="head" header>
          <Cell id="h1">Description</Cell>
          <Cell id="h2">Amount</Cell>
          <Cell id="h3">Note</Cell>
        </Row>
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
      is.within(table.gridTwips[0], FIXED_MM[0] * TWIPS_PER_MM, 1, "the first is 60mm wide");
      is.within(table.gridTwips[1], FIXED_MM[1] * TWIPS_PER_MM, 1, "the second is 40mm");
      is.within(table.gridTwips[2], AUTO_MM * TWIPS_PER_MM, 1, "and the third takes what is left");

      is.equal(
        table.width?.size,
        table.gridTwips.reduce((total, width) => total + width, 0),
        "the table declares itself exactly as wide as its columns",
      );
      is.equal(table.width?.type, "dxa", "in twips, not as a percentage of something");

      // The one that costs the most when it is missing: without it the grid
      // above is a suggestion Word is free to ignore, and it does.
      is.equal(table.layout, "fixed", "and declares a fixed layout, so the grid is not a suggestion");
    },

    preview: (b, is) => {
      is.within(b.table(0).w, b.mm(COLUMN_MM), "1mm", "the table fills the text column");
      is.within(b.cell("Consultancy").w, b.mm(FIXED_MM[0]), "1mm", "the first column draws 60mm wide");
      is.within(b.cell("1,250.00").w, b.mm(FIXED_MM[1]), "1mm", "the second draws 40mm");
      is.within(b.cell("Fixed fee").w, b.mm(AUTO_MM), "1mm", "the third draws what is left");
      is.within(b.cell("Consultancy").x, 0, "0.2mm", "and the table starts at the column's left edge");
    },

    word: (c, is) => {
      is.equal(c.table(0).rowCount, 2, "Word reads two rows");
      is.equal(c.table(0).columnCount, 3, "and three columns");
      is.equal(c.table(0).allowAutoFit, false, "and is not free to resize them");

      is.within(c.cell("Consultancy").width, c.mm(FIXED_MM[0]), "1mm", "Word makes the first 60mm");
      is.within(c.cell("1,250.00").width, c.mm(FIXED_MM[1]), "1mm", "the second 40mm");
      is.within(c.cell("Fixed fee").width, c.mm(AUTO_MM), "1mm", "and the third what is left");

      // Through the cells rather than through the table. Word reports
      // PreferredWidth as undefined for a fixed-layout table — the widths are
      // the columns' now, not a preference the table expresses — so asking the
      // table would measure the reading rather than the document.
      is.within(
        c.table(0).cells.slice(0, 3).reduce((total, cell) => total + (cell.width ?? 0), 0),
        c.mm(COLUMN_MM),
        "1mm",
        "which between them fill the text column exactly",
      );
    },

    /**
     * The claim the other three tiers cannot make between them: that the
     * columns a reader saw on screen are the columns Word prints.
     */
    parity: (p, is) => {
      is.within(p.previewCellX("Consultancy"), p.wordCellX("Consultancy"), "1mm", "the first column starts where Word starts it");
      is.within(p.previewCellY("Consultancy"), p.wordCellY("Consultancy"), "1mm", "and on the line Word puts it on");
      is.within(p.previewCellX("1,250.00"), p.wordCellX("1,250.00"), "1mm", "the second column starts where Word starts it");
      is.within(p.previewCellX("Fixed fee"), p.wordCellX("Fixed fee"), "1mm", "and so does the third");
      is.within(p.previewCellWidth("1,250.00"), p.wordCellWidth("1,250.00"), "1mm", "the money column is as wide on screen as in Word");
    },
  },
});
