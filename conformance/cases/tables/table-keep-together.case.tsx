import { Cell, Document, Paragraph, Row, Table, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * A table kept whole: it moves to the next page rather than breaking between rows.
 *
 * Without `keepTogether`, Word breaks a short table wherever the page ends, so
 * a totals box is left with its first lines at the foot of one page and its
 * last at the head of the next. Every row is `w:cantSplit` and every paragraph
 * but the last row's is `w:keepNext`. The filler count is tuned so the table
 * would straddle the page break if it were allowed to.
 *
 * No preview or parity assertion: the paginator does not consult keeps, so one
 * would only pass by accident.
 */
const FILLER = Array.from({ length: 27 }, (_, index) => index + 1);
const TOTALS = [1, 2, 3, 4, 5, 6];

export default defineCase({
  id: "tables/table-keep-together",
  feature: "table.keepTogether",
  title: "A table kept together moves to the next page whole rather than breaking between rows",
  word: "Table rows → Allow row to break across pages unchecked, paragraphs → Keep with next (w:cantSplit + w:keepNext)",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="table-keep-together" title="Table keep together">
      {FILLER.map((line) => <Paragraph id={`f${line}`}>{`Filler line ${line}`}</Paragraph>)}
      <Table id="totals" keepTogether columns={[{ width: "auto" }, { width: 30, align: "right" }]}>
        {TOTALS.map((line) => (
          <Row id={`t${line}`}>
            <Cell id={`t${line}c1`}>{`Totals line ${line}`}</Cell>
            <Cell id={`t${line}c2`}>10.00</Cell>
          </Row>
        ))}
      </Table>
    </Document>
  ),

  regions: [
    { id: "first-total", anchor: "Totals line 1" },
    { id: "last-total", anchor: "Totals line 6" },
  ],

  expect: {
    ooxml: (a, is) => {
      const table = a.table(0);

      for (let i = 0; i < table.rowCount; i++) {
        is.equal(table.row(i).cantSplit, true, `row ${i} declares w:cantSplit`);
      }
      is.equal(a.para("Totals line 1").keepNext, true, "a row but the last keeps with the next");
      is.equal(a.para("Totals line 6").keepNext, null, "the last row keeps with nothing");
      is.equal(a.para("Filler line 1").keepNext, null, "ordinary prose is not kept");
    },

    word: (c, is) => {
      is.equal(c.para("Totals line 1").keepWithNext, true, "Word reads keep with next");
      is.equal(c.table(0).row(0).allowBreakAcrossPages, false, "and the row as unbreakable");
      is.equal(
        c.cell("Totals line 1").page,
        c.cell("Totals line 6").page,
        "the whole table is on one page",
      );
      is.equal(c.cell("Totals line 1").page, 2, "and it moved whole to page two");
    },
  },
});
