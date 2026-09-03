import { Cell, Document, Paragraph, Row, Table, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * A table longer than a page, broken between two of its rows.
 *
 * `preview/content-pagination` established that the paginator matches Word
 * exactly on ordinary paragraphs. This is the same question asked of a table,
 * and when it was first asked the answer was two separate failures:
 *
 * 1. **A table could not be broken at all.** The paginator moves whole blocks,
 *    and a table is one block, so a table taller than the page had nowhere to
 *    go.
 * 2. **So the page grew instead.** docx-preview writes the page height as a
 *    `min-height`, which a block too tall to move simply pushes past —
 *    silently. What a reader saw was one sheet 593mm deep standing above a
 *    second of ordinary A4. A sheet of paper is not a thing that stretches,
 *    and a preview showing one that does is showing something the document
 *    cannot become.
 *
 * `splitTable` fixes both: a table is now broken at the last row that fits and
 * the rest carried onto the next sheet as a shallow clone, with the
 * `<colgroup>` copied so the columns line up across the seam. The second
 * failure goes with the first — nothing is left overrunning, so nothing
 * stretches a page.
 *
 * **The seam sits within one row of Word's, and that residual is measured
 * rather than hidden.** Word breaks this table 28 / 28 / 4 and the preview
 * breaks it 27 / 27 / 6: both take three pages and agree about where the first
 * row, the last row and the paragraph after them land, and the preview is one
 * row conservative at each seam. That is the leading drift `text/line-height`
 * records — a fraction of a line per line, which over twenty-seven rows comes
 * to one row — rather than anything about tables. So this case asserts the
 * page counts and the placements the two engines agree on, and takes care to
 * ask about a row either side of a seam rather than one on it.
 */
const LINES = Array.from({ length: 60 }, (_, index) => index + 1);

export default defineCase({
  id: "preview/table-pagination",
  feature: "preview.tablePagination",
  title: "The preview breaks a table that runs past the foot of the page",
  word: "Word breaks a table between two rows; so, now, does the preview",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="table-pagination" title="Table pagination">
      <Table id="ledger" columns={[{ width: "auto" }, { width: 30, align: "right" }]}>
        {LINES.map((line) => (
          <Row id={`r${line}`}>
            <Cell id={`r${line}c1`}>{`Ledger line ${line}`}</Cell>
            <Cell id={`r${line}c2`}>10.00</Cell>
          </Row>
        ))}
      </Table>
      <Paragraph id="after">Everything above this line is one table.</Paragraph>
    </Document>
  ),

  regions: [
    { id: "first", anchor: "Ledger line 1" },
    { id: "last", anchor: "Ledger line 60" },
    { id: "after", anchor: "Everything above this line" },
  ],

  expect: {
    /**
     * Nothing is wrong with the file, and nothing about it changed. It is one
     * table of sixty rows, and both engines lay it out over three pages from
     * exactly these bytes — which is what makes this a question about reading
     * rather than about packing.
     */
    ooxml: (a, is) => {
      is.equal(a.table(0).rowCount, 60, "one table, sixty rows");
      is.equal(a.paragraphCount, 1, "and one paragraph after it");
    },

    preview: (b, is) => {
      is.equal(b.pageCount(), 3, "the preview draws three pages");
      is.equal(b.tables.length, 3, "with the table broken into a piece on each");

      // The fix for the second failure, and the one worth keeping an assertion
      // on: a page that could not hold its contents used to grow to fit them.
      is.within(b.sections[0].page.h, b.mm(297), "1mm", "the first page is a sheet of A4");
      is.within(b.sections[1].page.h, b.mm(297), "1mm", "so is the second");
      is.within(b.sections[2].page.h, b.mm(297), "1mm", "and so is the last");

      // Every piece within the room its page has, which is what "broken at the
      // last row that fits" means once it is measured rather than asserted.
      is.less(b.table(0).h, b.mm(257), "the first piece fits the height of the text column");
      is.less(b.table(1).h, b.mm(257), "and so does the second");

      // The columns survive the seam. The widths live in the `<colgroup>`, and
      // a carried table without one would be laid out around its own contents
      // — the same autofit `tables/column-widths` is about, arriving by
      // another road.
      is.within(b.table(1).w, b.table(0).w, "0.5mm", "the second piece is as wide as the first");
      is.within(
        b.table(1).rows[0].cells[0].w,
        b.table(0).rows[0].cells[0].w,
        "0.5mm",
        "and its columns line up with them",
      );

      is.equal(b.cell("Ledger line 1").pageIndex, 0, "the first row is on page one");
      is.equal(b.cell("Ledger line 60").pageIndex, 2, "the last on page three");
      is.equal(b.para("Everything above this line").pageIndex, 2, "with the paragraph after it");
    },

    word: (c, is) => {
      is.equal(c.pageCount(), 3, "Word prints three pages from the same bytes");
      is.equal(c.cell("Ledger line 1").page, 1, "the first rows on page one");
      is.equal(c.cell("Ledger line 60").page, 3, "and the last on page three");
      is.equal(c.para("Everything above this line").page, 3, "with the paragraph after them");
    },

    /**
     * The number this case exists to produce: whether the two engines agree
     * about where the document ends, and about which page a row is on.
     */
    parity: (p, is) => {
      is.equal(p.previewPages(), p.wordPages(), "the preview breaks the document into as many pages as Word");
      is.equal(p.previewCellPage("Ledger line 1"), p.wordCellPage("Ledger line 1"), "and puts the first row on the page Word puts it on");
      // Line 29 rather than line 28: 28 is the row Word ends its first page on
      // and the preview opens its second with, which is the single row of seam
      // the two engines do not share.
      is.equal(p.previewCellPage("Ledger line 29"), p.wordCellPage("Ledger line 29"), "the row after the first seam on the page Word puts it on");
      is.equal(p.previewCellPage("Ledger line 60"), p.wordCellPage("Ledger line 60"), "and the last row too");
      is.equal(p.previewPage("Everything above this line"), p.wordPage("Everything above this line"), "with the paragraph after the table");
    },
  },
});
