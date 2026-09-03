import { Cell, Document, Row, Table, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * A header row that repeats at the top of every page the table runs onto.
 *
 * This is a thing only the renderer can do. A second header written into the
 * body would be a row of text saying the same words in the wrong place — it
 * would move when the table grew, print twice when the table fitted, and stay
 * behind when a row above it was dropped. `w:tblHeader` hands the problem to
 * whatever is laying the pages out, which is the only thing that knows where
 * the pages end.
 *
 * Which is why Word can only be asked half the question. It repeats a header
 * in the *layout*, not in the document, so the words appear once however many
 * pages they print on — there is no second "Description" to count. What is
 * assertable of Word is that it marks the row as a heading and that the table
 * really does cross a page, which is the pair that makes the repeat happen.
 * The preview is the tier where the repeat itself can be counted, because
 * there the second heading is a second row in the DOM.
 *
 * **The preview could do none of this when the case was written, and getting
 * it to took three pieces.** The paginator could not break a table, so there
 * was no second page to repeat onto — `preview/table-pagination` is where that
 * is measured. `w:tblHeader` is a fact docx-preview neither reads nor records,
 * so even given the page nothing knew which row was the heading: `readPackedTables`
 * reads it from the file and `settleDocxPreview` moves those rows into a
 * `<thead>`, which is what the element already means. And `splitTable` copies
 * the `<thead>` onto the carried table rather than moving it, because a
 * heading repeats where a row moves.
 *
 * Only the rows a table *opens* with are its heading. A `header` row further
 * down would print a subtotal at the top of every page above the figures it
 * adds up, so the packer counts the leading run and stops at the first body
 * row — asserted here by giving the last row `header` and watching it be
 * ignored.
 */
const LINES = Array.from({ length: 44 }, (_, index) => index + 1);

export default defineCase({
  id: "tables/header-row-repeat",
  feature: "table.headerRow",
  title: "A header row repeats on every page the table runs onto",
  word: "Table Properties → Row → Repeat as header row (w:tblHeader)",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="header-row-repeat" title="Header row repeat">
      <Table id="charges" columns={[{ width: "auto" }, { width: 30, align: "right" }]}>
        <Row id="head" header>
          <Cell id="h1">Description</Cell>
          <Cell id="h2">Amount</Cell>
        </Row>
        {LINES.map((line) => (
          <Row id={`r${line}`}>
            <Cell id={`r${line}c1`}>{`Charge line ${line}`}</Cell>
            <Cell id={`r${line}c2`}>10.00</Cell>
          </Row>
        ))}
        <Row id="total" header>
          <Cell id="t1">Total due</Cell>
          <Cell id="t2">440.00</Cell>
        </Row>
      </Table>
    </Document>
  ),

  regions: [
    { id: "head", anchor: "Description" },
    { id: "first", anchor: "Charge line 1" },
    { id: "last", anchor: "Charge line 44" },
  ],

  expect: {
    ooxml: (a, is) => {
      const table = a.table(0);

      is.equal(table.rowCount, 46, "the table is a header, forty-four lines and a total");
      is.equal(table.row(0).tblHeader, true, "the row it opens with declares w:tblHeader");
      is.equal(table.row(1).tblHeader, false, "the first body row explicitly does not");
      is.equal(
        table.row(45).tblHeader,
        false,
        "and neither does the totals row, header or not — only a leading run repeats",
      );

      // The words are written once. Everything below is about a second copy
      // that only exists once something has laid the pages out.
      is.equal(a.cells("Description").length, 1, "and the heading is written into the file once");
    },

    preview: (b, is) => {
      is.equal(b.pageCount(), 2, "the preview breaks the table onto a second page");
      is.equal(b.tables.length, 2, "with a piece of the table on each");

      // The assertion the case is for, and the only tier that can make it: a
      // second heading, drawn at the top of the second piece.
      is.equal(b.cells("Description").length, 2, "and draws the header row again at the top of it");
      is.equal(b.table(1).rows[0].cells[0].text, "Description", "which is the first row of the second piece");
      is.equal(b.cell("Charge line 1").pageIndex, 0, "the first charge is on page one");
      is.equal(b.cell("Charge line 44").pageIndex, 1, "and the last on page two");

      // Not the totals row, which is a `header` row the packer declined to
      // treat as one. If the leading run were miscounted it would repeat here.
      is.equal(b.cells("Total due").length, 1, "and the totals row is drawn once, header or not");
    },

    word: (c, is) => {
      is.equal(c.pageCount(), 2, "the table spills onto a second page");
      is.equal(c.table(0).row(0).headingFormat, true, "Word reads the opening row as a heading");
      is.equal(c.table(0).row(1).headingFormat, false, "and the first body row as an ordinary one");
      is.equal(c.table(0).row(45).headingFormat, false, "and the totals row as one too");

      // The pair that makes the repeat happen: a heading row, and a table with
      // a second page to repeat it onto.
      is.equal(c.cell("Charge line 1").page, 1, "the first charge is on page one");
      is.equal(c.cell("Charge line 44").page, 2, "and the last on page two");
    },

    /**
     * Which page each end of the table landed on, on screen against in Word.
     * The seam itself is within a row of Word's for the reason
     * `preview/table-pagination` sets out, so this asks about the rows either
     * side of it rather than the one on it.
     */
    parity: (p, is) => {
      is.equal(p.previewPages(), p.wordPages(), "the preview breaks the table onto as many pages as Word");
      is.equal(p.previewCellPage("Charge line 1"), p.wordCellPage("Charge line 1"), "the first charge on the page Word puts it on");
      is.equal(p.previewCellPage("Charge line 44"), p.wordCellPage("Charge line 44"), "and the last one too");
    },
  },
});
