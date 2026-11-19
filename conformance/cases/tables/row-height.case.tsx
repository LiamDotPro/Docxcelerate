import { Cell, Document, Row, Table, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { TWIPS_PER_MM, withBlocks } from "../_support/style.ts";

/**
 * A row given a height of its own, rather than the height of its contents.
 *
 * A band of a stated depth — a coloured strip across a letterhead, a signature
 * box, a row of even squares — is `w:trHeight`, which is what Word offers in
 * its own dialog and what `height` on a `<Row>` writes. It is not a block's
 * `heightPt`, which sets the *leading* of a paragraph and stops working the
 * moment the cell holds two lines or a picture.
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
  claim: "supported",

  style: withBlocks({
    band: { fill: "1F2933", color: "FFFFFF" },
  }),

  document: template(
    <Document id="row-height" title="Row height">
      <Table id="bands" columns={[{ width: "auto" }]}>
        <Row id="tall" height={HEIGHT_MM * (72 / 25.4)}>
          <Cell id="c1" variant="band">A band twenty millimetres deep</Cell>
        </Row>
        {/* docx-preview grows an exact row whose contents overflow, where Word clips it; deliberate, not a bug. */}
        <Row id="fixed" height={HEIGHT_MM * (72 / 25.4)} heightRule="exact">
          <Cell id="cf">A fixed row, the same depth</Cell>
        </Row>
        <Row id="ordinary">
          <Cell id="c2">An ordinary row, as tall as its one line of text</Cell>
        </Row>
      </Table>
    </Document>
  ),

  regions: [
    { id: "band", anchor: "A band twenty millimetres deep" },
    { id: "fixed", anchor: "A fixed row" },
    { id: "ordinary", anchor: "An ordinary row" },
  ],

  expect: {
    ooxml: (a, is) => {
      const table = a.table(0);

      is.equal(table.rowCount, 3, "three rows, two of which want a height");
      is.equal(table.row(2).heightTwips, null, "the ordinary row declares none, correctly");

      is.within(
        table.row(0).heightTwips,
        HEIGHT_MM * TWIPS_PER_MM,
        1,
        "the band declares a height of 20mm",
      );
      is.equal(table.row(0).heightRule, "atLeast", "as a floor rather than as a ceiling that clips it");
      is.within(table.row(1).heightTwips, HEIGHT_MM * TWIPS_PER_MM, 1, "the fixed row declares the same height");
      is.equal(table.row(1).heightRule, "exact", "as a fixed depth");
    },

    preview: (b, is) => {
      is.within(b.row(0, 1).h, b.mm(HEIGHT_MM), "1mm", "the fixed row is as deep");
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
      is.equal(c.table(0).row(1).heightRule, "exact", "the fixed row is exact");
      is.equal(c.table(0).row(2).heightRule, "auto", "leaving the ordinary row to its contents");
    },
  },
});
