import { Cell, Document, Paragraph, Row, Table, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * A row that stays whole: it moves to the next page rather than breaking.
 *
 * Without `keepTogether`, Word breaks a row wherever the page ends, so a
 * schedule item of several lines is left with its top at the foot of one page
 * and the rest at the head of the next. `w:cantSplit` hands the row to the
 * layout whole. The filler count is tuned so that the tall row would straddle
 * the page break if it were allowed to: about twenty-five single lines fill
 * most of the first page, and the tall row needs more room than is left.
 *
 * The preview only ever breaks a table between rows, so it draws every row
 * whole already; this case pins that the two agree on where the row lands.
 */
const LINES = Array.from({ length: 25 }, (_, index) => index + 1);

export default defineCase({
  id: "tables/row-keep-together",
  feature: "table.rowCantSplit",
  title: "A row kept whole moves to the next page rather than breaking",
  word: "Table Properties → Row → Allow row to break across pages, unchecked (w:cantSplit)",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="row-keep-together" title="Row keep together">
      <Table id="charges" columns={[{ width: "auto" }, { width: 30, align: "right" }]}>
        {LINES.map((line) => (
          <Row id={`r${line}`}>
            <Cell id={`r${line}c1`}>{`Charge line ${line}`}</Cell>
            <Cell id={`r${line}c2`}>10.00</Cell>
          </Row>
        ))}
        <Row id="schedule" keepTogether>
          <Cell id="s1">
            <Paragraph id="s1p">{"Schedule item\nLine two\nLine three\nLine four\nLine five\nLine six\nLine seven\nLine eight"}</Paragraph>
          </Cell>
          <Cell id="s2">250.00</Cell>
        </Row>
      </Table>
    </Document>
  ),

  regions: [
    { id: "last-charge", anchor: "Charge line 25" },
    { id: "schedule", anchor: "Schedule item" },
  ],

  expect: {
    ooxml: (a, is) => {
      const table = a.table(0);
      const LAST = table.rowCount - 1;

      is.equal(table.row(LAST).cantSplit, true, "the schedule row declares w:cantSplit");
      is.equal(table.row(0).cantSplit, null, "an unmarked row writes nothing");
    },

    preview: (b, is) => {
      is.equal(b.cell("Charge line 25").pageIndex, 0, "the last charge is on page one");
      is.equal(b.cell("Schedule item").pageIndex, 1, "and the whole schedule row is on page two");
    },

    word: (c, is) => {
      const LAST = c.table(0).rowCount - 1;

      is.equal(c.table(0).row(LAST).allowBreakAcrossPages, false, "Word reads the row as unbreakable");
      is.equal(c.table(0).row(0).allowBreakAcrossPages, true, "and an ordinary row as breakable");
      is.equal(c.cell("Charge line 25").page, 1, "the last charge is on page one");
      is.equal(c.cell("Schedule item").page, 2, "and the schedule row moved whole to page two");
    },

    parity: (p, is) => {
      is.equal(
        p.previewCellPage("Schedule item"),
        p.wordCellPage("Schedule item"),
        "the preview puts the row on the page Word does",
      );
    },
  },
});
