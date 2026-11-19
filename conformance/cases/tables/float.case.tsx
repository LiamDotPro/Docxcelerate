import { Cell, Document, Paragraph, Row, Table, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * Text that runs around a table.
 *
 * **Not expressible.** A table always occupies the full band of the page it is
 * on: the packer writes no `w:tblpPr`, so there is no way to set one beside a
 * paragraph. Alignment is a separate, supported thing — see `tables/align`.
 *
 * Floating is not small: a `w:tblpPr` takes the table out of the flow entirely
 * and makes the paragraphs around it wrap, and a preview that has to reproduce
 * *that* is a preview doing its own line breaking. This case is where that
 * question is written down rather than discovered halfway through.
 */
export default defineCase({
  id: "tables/float",
  feature: "table.float",
  title: "A table text runs around",
  word: "Text wrapping (w:tblpPr)",
  claim: "unsupported",

  style: caseStyle,

  document: template(
    <Document id="float" title="Table alignment and float">
      <Paragraph id="before">A paragraph above the table, which is where every paragraph goes.</Paragraph>

      <Table id="narrow" columns={[{ width: 60 }]}>
        <Row id="head" header>
          <Cell id="h1">A narrow table</Cell>
        </Row>
        <Row id="r1">
          <Cell id="c1">Sixty millimetres wide</Cell>
        </Row>
      </Table>

      <Paragraph id="after">A paragraph below it, which is the only other place one can go.</Paragraph>
    </Document>
  ),

  regions: [
    { id: "table", anchor: "A narrow table" },
    { id: "after", anchor: "A paragraph below it" },
  ],

  expect: {
    ooxml: (a, is) => {
      const table = a.table(0);

      // What is true today, so the gap has a floor under it.
      is.equal(table.gridTwips.length, 1, "the table is one narrow column");
      is.equal(table.indent, null, "sitting against the left margin");

      // The gap.
      is.equal(table.floating, true, "and declares a position for the text to run around");
    },

    word: (c, is) => {
      is.equal(c.table(0).wrapAroundText, true, "and the text as running around it");
    },
  },
});
