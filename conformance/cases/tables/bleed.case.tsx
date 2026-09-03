import { Cell, Document, Row, Table, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle, COLUMN_MM, TWIPS_PER_MM, withBlocks } from "../_support/style.ts";

/**
 * A table pulled out past the margins to the edge of the sheet.
 *
 * What a coloured band across the top of a letterhead is, and a footer bar.
 * Word measures a table from the text column's left edge, so reaching the
 * paper means indenting *backwards* by exactly the width of the margin — a
 * negative `w:tblInd` — and widening the table by both margins to match.
 * A bar that stopped at the margin would be a bar with a white gutter either
 * side of it, which is not a bar.
 *
 * This is also the case behind one of the preview's fixes, and the reason it
 * is worth a tier of its own. docx-preview reads a table's indent with
 * `parseIndentation`, which looks for a `w:left` attribute — and `w:tblInd`
 * does not carry one, its value is in `w:w`. So the indent was parsed and then
 * dropped, and a table the file says reaches the paper's edge stopped at the
 * margin on screen while Word printed it to the edge. `applyTableIndents` puts
 * it back, from the file rather than from the theme, and the parity tier is
 * what says it put it back in the right place.
 */
const MARGIN_MM = caseStyle.page.margins.leftMm;
const PAGE_MM = 210;

export default defineCase({
  id: "tables/bleed",
  feature: "table.bleed",
  title: "A table that reaches the edge of the paper",
  word: "Table Properties → Indent from left, negative (w:tblInd)",
  claim: "supported",

  style: withBlocks({
    /** A band of colour across the sheet, edge to edge. */
    bar: { bleed: true, fill: "1F2933", color: "FFFFFF", paddingPt: 8 },
  }),

  document: template(
    <Document id="bleed" title="Bleeding table">
      <Table id="band" variant="bar" columns={[{ width: "auto" }]}>
        <Row id="r1">
          <Cell id="c1">A band that reaches both edges of the sheet</Cell>
        </Row>
      </Table>

      <Table id="inside" columns={[{ width: "auto" }]}>
        <Row id="r2">
          <Cell id="c2">An ordinary table, which stops at the margin</Cell>
        </Row>
      </Table>
    </Document>
  ),

  regions: [
    { id: "band", anchor: "A band that reaches" },
    { id: "inside", anchor: "An ordinary table" },
  ],

  expect: {
    ooxml: (a, is) => {
      const band = a.table(0);

      is.within(
        band.indent?.size,
        -MARGIN_MM * TWIPS_PER_MM,
        1,
        "the band is indented backwards by exactly the margin",
      );
      is.equal(band.indent?.type, "dxa", "in twips, which is the unit the margin is in");
      is.within(
        band.width?.size,
        PAGE_MM * TWIPS_PER_MM,
        2,
        "and is as wide as the whole sheet, not as the text column",
      );

      is.equal(a.table(1).indent, null, "the ordinary table declares no indent");
      is.within(a.table(1).width?.size, COLUMN_MM * TWIPS_PER_MM, 2, "and is the width of the text column");
    },

    preview: (b, is) => {
      is.within(b.table(0).x, -b.mm(MARGIN_MM), "0.5mm", "the preview pulls the band back past the margin");
      is.within(b.table(0).w, b.mm(PAGE_MM), "1mm", "and draws it the full width of the sheet");
      is.within(b.table(1).x, 0, "0.5mm", "while the ordinary table starts at the text column");
      is.within(b.table(1).w, b.mm(COLUMN_MM), "1mm", "and is as wide as it");
    },

    word: (c, is) => {
      is.within(c.table(0).leftIndent, -c.mm(MARGIN_MM), "0.5mm", "Word indents the band back past the margin");
      is.within(c.table(0).preferredWidth, c.mm(PAGE_MM), "1mm", "and makes it the width of the sheet");
      is.within(c.table(1).leftIndent, 0, "0.5mm", "and leaves the ordinary table on the margin");
    },

    /**
     * Where the band's own left edge lands, on screen against in Word. This is
     * the assertion `applyTableIndents` exists to satisfy — before it, the
     * preview drew this table 20mm to the right of where Word prints it.
     */
    parity: (p, is) => {
      is.within(p.previewCellX("A band that reaches"), p.wordCellX("A band that reaches"), "1mm", "the band's words start where Word starts them");
      is.within(p.previewCellX("An ordinary table"), p.wordCellX("An ordinary table"), "1mm", "and the ordinary table's do too");
      is.within(p.previewTableX(0), -MARGIN_MM, "1mm", "with the band's left edge on the paper's edge");
    },
  },
});
