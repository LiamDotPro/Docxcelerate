import { Cell, Document, Paragraph, Row, Table, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { TWIPS_PER_MM, withBlocks } from "../_support/style.ts";

/**
 * A paragraph inside a cell keeps its own layout: a dotted leader in a
 * signature cell, a line ranged right in a left column, a ruled quote inset.
 *
 * Before, a cell rebuilt every paragraph from the column alone, so each of
 * these needed a table of its own or could not be written. The cell's own block
 * still stays on the cell — only the paragraph's variant draws on the paragraph.
 *
 * **One column, on purpose**, for the reason `tables/cell-align` gives: Word
 * autofits a multi-column grid around its contents, so nothing measured here
 * would be the layout rather than the width gap.
 */
export default defineCase({
  id: "tables/cell-paragraph",
  feature: "table.cellParagraph",
  title: "A paragraph in a cell keeps its tab stops, alignment and inset",
  word: "Paragraph → Tabs, Alignment, Indentation, Borders and Shading, inside a table cell",
  claim: "supported",

  style: withBlocks({
    signLine: { tabStopsMm: [{ at: 120, align: "right", leader: "dot" }] },
    inset: {
      indentMm: 10,
      indentRightMm: 10,
      border: "2F5FBD",
      borderSides: ["left"],
      borderWidthPt: 2,
      paddingPt: 6,
      fill: "EEF2FF",
    },
  }),

  document: template(
    <Document id="cell-paragraph" title="Cell paragraphs">
      <Table id="panel" columns={[{ width: "auto" }]}>
        <Row id="a">
          <Cell id="a1">
            <Paragraph id="sign" variant="signLine">{"Signed for the supplier\t19 November 2026"}</Paragraph>
          </Cell>
        </Row>
        <Row id="b">
          <Cell id="b1">
            <Paragraph id="note" align="right">Ranged right in a left column</Paragraph>
          </Cell>
        </Row>
        <Row id="c">
          <Cell id="c1">
            <Paragraph id="quote" variant="inset">Quoted inside a card</Paragraph>
          </Cell>
        </Row>
      </Table>
    </Document>
  ),

  regions: [
    { id: "sign", anchor: "Signed for the supplier" },
    { id: "note", anchor: "Ranged right in a left column" },
    { id: "quote", anchor: "Quoted inside a card" },
  ],

  expect: {
    ooxml: (a, is) => {
      const stop = a.cell("Signed for the supplier").paragraphs[0]?.tabs[0];

      is.within(stop?.pos, 120 * TWIPS_PER_MM, 2, "the signature stop is 120mm in");
      is.equal(stop?.val, "right", "as a right stop");
      is.equal(stop?.leader, "dot", "with a dot leader");

      is.equal(a.cell("Ranged right in a left column").paragraphs[0]?.jc, "right", "the paragraph's own align reaches it");

      const quote = a.cell("Quoted inside a card").paragraphs[0];

      is.within(quote?.ind?.left, 567, 2, "the variant's indent is written");
      is.equal(quote?.shd, "EEF2FF", "and its fill");
      is.equal(quote?.pBdr?.left === null || quote?.pBdr === null, false, "and its left rule");
    },

    preview: (b, is) => {
      const signed = b.cell("Signed for the supplier");

      is.within(
        b.cellTextRight("Signed for the supplier"),
        signed.x + signed.paddingLeft + b.mm(120),
        "2mm",
        "the date ends on the stop, measured from the cell's inner left edge",
      );

      const note = b.cell("Ranged right in a left column");

      is.within(
        b.cellTextRight("Ranged right in a left column"),
        note.x + note.w - note.paddingRight,
        "0.5mm",
        "the right-ranged note ends at the cell's inner right edge",
      );

      const quote = b.cell("Quoted inside a card");

      is.within(
        b.cellTextLeft("Quoted inside a card"),
        quote.x + quote.paddingLeft + b.mm(10),
        "1mm",
        "the inset text starts 10mm inside the cell",
      );
    },

    word: (c, is) => {
      const stops = c.customTabStops("Signed for the supplier");

      is.within(stops[0]?.position, c.mm(120), "0.5mm", "Word reads the stop 120mm in");
      is.equal(stops[0]?.leader, "dot", "with its dot leader");
      is.equal(c.cell("Ranged right in a left column").alignment, "right", "and the note ranged right");
    },

    parity: (p, is) => {
      is.within(
        p.previewRight("Signed for the supplier"),
        p.wordRight("Signed for the supplier"),
        "2mm",
        "the dated end of the signed line is where Word puts it",
      );
    },
  },
});
