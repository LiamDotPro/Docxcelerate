import { Cell, Document, Row, Table, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle, TWIPS_PER_MM, withBlocks } from "../_support/style.ts";

/**
 * The same two tables on a turned page.
 *
 * `w:pgSz` has always turned, but the widths worked out from the page used the
 * upright sheet, so on landscape A4 an auto table stopped 87mm short of the
 * right margin. The band reaches both edges of the 297mm sheet and the
 * ordinary table fills the 257mm column between the margins.
 */
const MARGIN_MM = caseStyle.page.margins.leftMm;
const PAGE_MM = 297;
const COLUMN_MM = PAGE_MM - 2 * MARGIN_MM;

export default defineCase({
  id: "page/orientation-landscape",
  feature: "page.orientation",
  title: "Tables on a landscape page",
  word: "Layout → Orientation → Landscape",
  claim: "supported",

  style: {
    ...withBlocks({
      /** A band of colour across the sheet, edge to edge. */
      bar: { bleed: true, fill: "1F2933", color: "FFFFFF", paddingPt: 8 },
    }),
    page: { ...caseStyle.page, orientation: "landscape" },
  },

  document: template(
    <Document id="landscape" title="Landscape tables">
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
      is.equal(a.section.orientation, "landscape", "the sheet is turned");
      is.within(a.section.widthTwips, PAGE_MM * TWIPS_PER_MM, 2, "so it is 297mm wide");
      is.within(a.section.heightTwips, 210 * TWIPS_PER_MM, 2, "and 210mm tall");
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
