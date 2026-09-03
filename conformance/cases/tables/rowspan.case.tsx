import { Cell, Document, Row, Table, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * A cell that runs down more than one row.
 *
 * **The model has no way to say it.** `TableCellNode` has a `span`, and it is
 * horizontal; there is no `rowSpan` beside it and nothing in the packer writes
 * `w:vMerge`. So a document that wants one quarter's label to stand beside its
 * three months has two choices today, and both of them are wrong in a way a
 * reader can see: repeat the label on every row, or write it once and leave
 * the cells under it blank — a column that looks like missing data rather than
 * like a heading.
 *
 * This case is the second of those, so the gap has a shape. It is what a
 * document has to do now, and what it should not have to do.
 *
 * `w:vMerge` is two halves rather than a count: `restart` on the cell that
 * opens the merge, `continue` on each one swallowed under it. That is why the
 * assertions name both — a packer that wrote only the first would produce a
 * table Word reads as a merge of one row, which is not a merge.
 *
 * Word's own tell is `Uniform`. A table with a vertical merge has rows whose
 * cells do not line up, so Word reports it as not uniform and refuses to hand
 * over its Rows collection at all — which is exactly why the Word probe walks
 * cells through the table's range instead. The reading survives the merge; it
 * simply finds no merge to survive.
 */
export default defineCase({
  id: "tables/rowspan",
  feature: "table.rowspan",
  title: "A cell that runs down more than one row",
  word: "Table → Merge Cells, vertically (w:vMerge restart / continue)",
  claim: "unsupported",

  style: caseStyle,

  document: template(
    <Document id="rowspan" title="Vertical merge">
      <Table id="quarters" columns={[{ width: 40 }, { width: 40 }, { width: "auto", align: "right" }]}>
        <Row id="head" header>
          <Cell id="h1">Quarter</Cell>
          <Cell id="h2">Month</Cell>
          <Cell id="h3">Amount</Cell>
        </Row>
        <Row id="jan">
          <Cell id="q1">Quarter one</Cell>
          <Cell id="m1">January</Cell>
          <Cell id="a1">1,000.00</Cell>
        </Row>
        <Row id="feb">
          <Cell id="q2"></Cell>
          <Cell id="m2">February</Cell>
          <Cell id="a2">1,100.00</Cell>
        </Row>
        <Row id="mar">
          <Cell id="q3"></Cell>
          <Cell id="m3">March</Cell>
          <Cell id="a3">1,200.00</Cell>
        </Row>
      </Table>
    </Document>
  ),

  regions: [
    { id: "label", anchor: "Quarter one" },
    { id: "first", anchor: "January" },
    { id: "last", anchor: "March" },
  ],

  expect: {
    ooxml: (a, is) => {
      const table = a.table(0);

      // What the file does say, so the case is a gap rather than a mystery.
      is.equal(table.rowCount, 4, "the table is a header and three months");
      is.equal(table.row(2).cellCount, 3, "and the row under the label still holds three cells");

      // The gap. Every one of these is null today.
      is.equal(a.cell("Quarter one").vMerge, "restart", "the label's cell opens a vertical merge");
      is.equal(table.row(2).cell(0).vMerge, "continue", "the cell below it continues that merge");
      is.equal(table.row(3).cell(0).vMerge, "continue", "and so does the one below that");
    },

    preview: (b, is) => {
      // A merged cell is one cell three rows tall. Today it is three cells,
      // two of them empty, and the difference is visible on the page.
      is.greater(
        b.cell("Quarter one").h,
        b.cell("January").h * 2,
        "the label's cell is drawn as tall as the rows it covers",
      );
      is.equal(b.row(0, 2).cells.length, 2, "so the row under it has one cell fewer");
    },

    word: (c, is) => {
      is.equal(c.table(0).uniform, false, "Word reads the table as merged rather than as a plain grid");
      is.greater(
        c.cell("Quarter one").height,
        (c.cell("January").height ?? 0) * 2,
        "with the label's cell as tall as the rows it covers",
      );
    },
  },
});
