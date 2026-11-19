import { Cell, Document, Paragraph, Row, Table, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * A table that sits somewhere other than against the left margin.
 *
 * Word's `w:jc` on a `w:tblPr` centres or ranges a table, and the text still
 * goes above and below it. docx-preview reads it as auto margins, so the
 * preview needs no rule of its own.
 *
 * The tables are deliberately narrow. A table the width of the text column is
 * centred, ranged left and ranged right in the same place, so a case that used
 * one would pass its alignment assertions by accident.
 */
export default defineCase({
  id: "tables/align",
  feature: "table.align",
  title: "A table centred on the page, and one ranged right",
  word: "Table Properties → Alignment (w:jc on w:tblPr)",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="align" title="Table alignment">
      <Table id="centred" align="center" columns={[{ width: 60 }]}>
        <Row id="r1">
          <Cell id="c1">A centred table</Cell>
        </Row>
      </Table>

      <Paragraph id="between">A paragraph between them, which stays where every paragraph goes.</Paragraph>

      <Table id="right" align="right" columns={[{ width: 60 }]}>
        <Row id="r2">
          <Cell id="c2">A table ranged right</Cell>
        </Row>
      </Table>
    </Document>
  ),

  regions: [
    { id: "centred", anchor: "A centred table" },
    { id: "right", anchor: "A table ranged right" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.equal(a.table(0).jc, "center", "the first table declares itself centred");
      is.equal(a.table(1).jc, "right", "the second declares itself ranged right");
      is.equal(a.table(0).gridTwips.length, 1, "each is one narrow column");
      is.equal(a.table(0).indent, null, "with no indent to fight the alignment");
    },

    preview: (b, is) => {
      const page = b.sections[0]?.content?.w ?? 0;
      is.within(b.table(0).x + b.table(0).w / 2, page / 2, "1mm", "the preview centres the first table on the text column");
      is.within(b.table(1).x + b.table(1).w, page, "1mm", "and ranges the second against the right edge");
    },

    word: (c, is) => {
      is.equal(c.table(0).alignment, "center", "Word reads the first table as centred");
      is.equal(c.table(1).alignment, "right", "and the second as ranged right");
    },

    parity: (p, is) => {
      is.within(p.previewCellX("A centred table"), p.wordCellX("A centred table"), "1mm", "the centred table starts where Word starts it");
      is.within(p.previewCellX("A table ranged right"), p.wordCellX("A table ranged right"), "1mm", "and the right-ranged one");
    },
  },
});
