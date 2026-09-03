import { Cell, Document, Row, Table, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { TWIPS_PER_MM, withBlocks } from "../_support/style.ts";

/**
 * A row given a height of its own, rather than the height of its contents.
 *
 * **The model has no way to say it.** `TableRowNode` carries a `header` and
 * its cells, and nothing else; there is no height on a row and the packer
 * writes no `w:trHeight`. A row is as tall as whatever is in it, always.
 *
 * That is fine until a document needs a band of a stated depth — a coloured
 * strip across a letterhead, a signature box, a row of even squares. Today the
 * nearest thing is a block's `heightPt`, which sets the *leading* of the
 * paragraph inside the cell to an exact number of points: it makes a row that
 * happens to come out about that tall, by way of a property that means
 * something else, and it stops working the moment the cell holds two lines or
 * a picture. `w:trHeight` says the thing directly and is what Word offers in
 * its own dialog.
 *
 * The rule matters as much as the number. `atLeast` is a floor — the row grows
 * if its contents need more — and `exact` is a ceiling that clips them. They
 * are the difference between a band that holds its depth and one that quietly
 * loses a descender, and a model that offered a height without a rule would be
 * offering the more dangerous of the two by accident.
 */
const HEIGHT_MM = 20;

export default defineCase({
  id: "tables/row-height",
  feature: "table.rowHeight",
  title: "A row given a height of its own",
  word: "Table Properties → Row → Specify height (w:trHeight, w:hRule)",
  claim: "unsupported",

  style: withBlocks({
    /** The nearest thing available: a cell whose leading is set in points. */
    band: { fill: "1F2933", color: "FFFFFF", heightPt: HEIGHT_MM * (72 / 25.4) },
  }),

  document: template(
    <Document id="row-height" title="Row height">
      <Table id="bands" columns={[{ width: "auto" }]}>
        <Row id="tall">
          <Cell id="c1" variant="band">A band twenty millimetres deep</Cell>
        </Row>
        <Row id="ordinary">
          <Cell id="c2">An ordinary row, as tall as its one line of text</Cell>
        </Row>
      </Table>
    </Document>
  ),

  regions: [
    { id: "band", anchor: "A band twenty millimetres deep" },
    { id: "ordinary", anchor: "An ordinary row" },
  ],

  expect: {
    ooxml: (a, is) => {
      const table = a.table(0);

      is.equal(table.rowCount, 2, "two rows, one of which wants a height");
      is.equal(table.row(1).heightTwips, null, "the ordinary row declares none, correctly");

      // The gap.
      is.within(
        table.row(0).heightTwips,
        HEIGHT_MM * TWIPS_PER_MM,
        1,
        "the band declares a height of 20mm",
      );
      is.equal(table.row(0).heightRule, "atLeast", "as a floor rather than as a ceiling that clips it");
    },

    preview: (b, is) => {
      is.within(
        b.row(0, 0).h,
        b.mm(HEIGHT_MM),
        "1mm",
        "the preview draws the band twenty millimetres deep",
      );
    },

    word: (c, is) => {
      is.within(c.table(0).row(0).height, c.mm(HEIGHT_MM), "1mm", "Word makes the band 20mm deep");
      is.equal(c.table(0).row(0).heightRule, "atLeast", "as a floor");
      is.equal(c.table(0).row(1).heightRule, "auto", "leaving the ordinary row to its contents");
    },
  },
});
