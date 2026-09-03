import { Cell, Document, Paragraph, Row, Table, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * A table that sits somewhere other than against the left margin, and text
 * that runs around it.
 *
 * **Neither is expressible.** A table is always left where it stands and
 * always occupies the full band of the page it is on: the packer writes no
 * `w:jc` on a table and no `w:tblpPr`, so there is no way to centre a narrow
 * table, to range one right, or to set one beside a paragraph.
 *
 * The two are one case because they are one decision, and it is worth saying
 * which way it should go. Table alignment is small and behaves — Word's `w:jc`
 * on a `w:tblPr` centres the table and the text still goes above and below it,
 * so the preview can reproduce it with a margin. Floating is not small: a
 * `w:tblpPr` takes the table out of the flow entirely and makes the paragraphs
 * around it wrap, and a preview that has to reproduce *that* is a preview
 * doing its own line breaking. The first belongs on the roadmap; the second is
 * a question rather than a task, and this case is where the difference is
 * written down rather than discovered halfway through.
 *
 * The narrow table is deliberately narrow. A table the width of the text
 * column is centred, ranged left and ranged right in the same place, so a case
 * that used one would pass its alignment assertions by accident.
 */
export default defineCase({
  id: "tables/float",
  feature: "table.float",
  title: "A table centred on the page, and text running around one",
  word: "Table Properties → Alignment, and Text wrapping (w:jc on w:tblPr, w:tblpPr)",
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
      is.equal(table.jc, "center", "the table declares itself centred on the page");
      is.equal(table.floating, true, "and declares a position for the text to run around");
    },

    preview: (b, is) => {
      const page = b.sections[0]?.content?.w ?? 0;
      is.within(
        b.table(0).x + b.table(0).w / 2,
        page / 2,
        "1mm",
        "the preview centres the table on the text column",
      );
    },

    word: (c, is) => {
      is.equal(c.table(0).alignment, "center", "Word reads the table as centred");
      is.equal(c.table(0).wrapAroundText, true, "and the text as running around it");
    },
  },
});
