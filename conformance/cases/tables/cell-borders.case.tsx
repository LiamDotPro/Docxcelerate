import { Cell, Document, Row, Table, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { TWIPS_PER_PT, withBlocks } from "../_support/style.ts";

/**
 * What a cell is drawn as: its fill, the edges it rules, and the room it
 * leaves around what it holds.
 *
 * A cell is the one place in the model where a box behaves the way a reader
 * expects a box to. `text/block-box` records the paragraph's version of this
 * and the asymmetry it carries — Word has no padding on a paragraph, only a
 * gap between a border and the text on the sides a border is actually drawn
 * on, so a tinted panel with no border has nothing to hold its words off its
 * own edge. A cell has `w:tcMar`, which is real room on all four sides whether
 * anything is ruled or not, and that difference is why a filled panel belongs
 * in a table and not in a paragraph.
 *
 * One column, so that Word's autofit (`tables/column-widths`) cannot move
 * anything sideways while the case is measuring how far in the words start.
 *
 * The unfilled row is not decoration. A table with a header gets a hairline
 * under each body row, in the palette's rule colour, and a row that draws its
 * own ground does not — a rule at the foot of a filled panel is a line cut
 * across it. So the plain row proves the hairline is there and the filled ones
 * prove it steps aside.
 *
 * **The padding used to be written twice, and this case is what found it.**
 * `blockBorders` builds one border object for both kinds of box, and it set
 * `w:space` from the block's padding — right for `w:pBdr`, where `w:space`
 * *is* the gap between the border and the text and the only room a paragraph
 * can have, and wrong for `w:tcBorders`, where it is a gap between the cell's
 * edge and the border and the room inside is `w:tcMar`'s job. So a panel with
 * 12pt of padding declared 12pt of `w:tcMar` and another 12pt of border space.
 *
 * Measured before the fix: docx-preview ignores a cell border's space and drew
 * the words 12pt below the cell's top edge; Word honoured it and drew them
 * 24pt below. Horizontally the two agreed — Word does not apply the space to a
 * left or right cell border — so it showed up as a panel sitting 3.5mm lower
 * in Word than on screen, and every row under it 3.5mm lower again. That is
 * the shape of the bug this suite exists for: three tiers green and the one
 * that compares them red. `blockBorders` now takes which element it is
 * drawing for, and the assertion below is what keeps the space off a cell.
 */
const PADDING_PT = 12;

export default defineCase({
  id: "tables/cell-borders",
  feature: "table.cellBox",
  title: "A cell's fill, the edges it rules, and the room inside it",
  word: "Table Properties → Borders and Shading, Cell margins (w:shd, w:tcBorders, w:tcMar)",
  claim: "supported",

  style: withBlocks({
    /** A tinted panel, ruled on all four sides and generously padded. */
    panel: {
      fill: "EEF2FF",
      border: "C7D2FE",
      borderWidthPt: 1,
      paddingPt: PADDING_PT,
    },
    /** Filled, with no rule anywhere — the case a paragraph cannot pad. */
    tinted: {
      fill: "FEF3C7",
      paddingPt: PADDING_PT,
      borderSides: [],
    },
  }),

  document: template(
    <Document id="cell-borders" title="Cell box">
      <Table id="boxes" columns={[{ width: "auto" }]}>
        <Row id="head" header>
          <Cell id="h1">Boxes</Cell>
        </Row>
        <Row id="plain">
          <Cell id="p1">A plain cell, ruled off from the next by a hairline.</Cell>
        </Row>
        <Row id="panel">
          <Cell id="pa1" variant="panel">A panel: tinted, ruled all round, and padded.</Cell>
        </Row>
        <Row id="tinted">
          <Cell id="t1" variant="tinted">Tinted, with no rule at all, and padded anyway.</Cell>
        </Row>
      </Table>
    </Document>
  ),

  regions: [
    { id: "plain", anchor: "A plain cell" },
    { id: "panel", anchor: "A panel: tinted" },
    { id: "tinted", anchor: "Tinted, with no rule" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.equal(a.cell("A plain cell").shd, null, "a plain cell is not filled");
      is.equal(
        a.cell("A plain cell").borders?.bottom?.color,
        "D1D5DB",
        "and carries the hairline that separates it from the next row",
      );

      is.equal(a.cell("A panel: tinted").shd, "EEF2FF", "the panel's fill is written as cell shading");
      is.equal(a.cell("A panel: tinted").borders?.top?.color, "C7D2FE", "and its rule is drawn on top");
      is.equal(a.cell("A panel: tinted").borders?.bottom?.color, "C7D2FE", "bottom");
      is.equal(a.cell("A panel: tinted").borders?.left?.color, "C7D2FE", "left");
      is.equal(a.cell("A panel: tinted").borders?.right?.color, "C7D2FE", "and right");

      is.equal(a.cell("Tinted, with no rule").shd, "FEF3C7", "the tinted cell is filled");
      is.equal(a.cell("Tinted, with no rule").borders, null, "and draws no edges at all");

      // The room inside, which is the half a paragraph cannot have. Written in
      // twips, so a padding stated in points is asserted in the file's unit.
      is.equal(
        a.cell("A panel: tinted").margins?.left,
        PADDING_PT * TWIPS_PER_PT,
        "the panel leaves 12pt between its left edge and its words",
      );
      is.equal(a.cell("A panel: tinted").margins?.top, PADDING_PT * TWIPS_PER_PT, "and above them");
      is.equal(
        a.cell("Tinted, with no rule").margins?.left,
        PADDING_PT * TWIPS_PER_PT,
        "and the tinted cell leaves it too, border or no border",
      );

      // The room inside a cell is w:tcMar's to state, and a cell border that
      // also carried the padding as its own w:space stated it twice.
      is.equal(
        a.cell("A panel: tinted").borders?.top?.spacePt,
        null,
        "and the panel's border claims no room of its own on top of that",
      );
    },

    preview: (b, is) => {
      is.equal(b.cell("A panel: tinted").background, b.hex("EEF2FF"), "the panel draws its fill");
      is.equal(b.cell("A panel: tinted").borderTopColor, b.hex("C7D2FE"), "and its rule");
      is.equal(b.cell("Tinted, with no rule").background, b.hex("FEF3C7"), "the tinted cell draws its fill");
      is.equal(b.cell("Tinted, with no rule").borderTopColor, null, "and no rule anywhere");
      is.equal(b.cell("A plain cell").borderBottomColor, b.hex("D1D5DB"), "the plain cell keeps its hairline");

      is.within(b.cell("A panel: tinted").paddingLeft, b.pt(PADDING_PT), "0.5mm", "with the words held 12pt off its left edge");
      is.within(b.cell("Tinted, with no rule").paddingLeft, b.pt(PADDING_PT), "0.5mm", "and the tinted cell's too");
    },

    word: (c, is) => {
      is.equal(c.cell("A panel: tinted").shading, "EEF2FF", "Word fills the panel");
      is.equal(c.cell("A panel: tinted").borders?.top, true, "and rules it on every side");
      is.equal(c.cell("A panel: tinted").borders?.left, true, "every side");
      is.equal(c.cell("Tinted, with no rule").shading, "FEF3C7", "Word fills the tinted cell");
      is.equal(c.cell("Tinted, with no rule").borders?.bottom, false, "and rules none of it");
      is.equal(c.cell("A plain cell").borders?.bottom, true, "the plain cell is ruled underneath");

      is.within(c.cell("A panel: tinted").padding?.left, PADDING_PT, 0.5, "Word leaves 12pt inside the panel");
      is.within(c.cell("Tinted, with no rule").padding?.left, PADDING_PT, 0.5, "and inside the tinted cell");
    },

    /**
     * Where the words sit inside the box, on screen against in Word.
     *
     * The assertion the case is really for. A cell whose text starts 12pt in
     * on screen and flush with the fill in Word is a preview that lied about
     * the one thing a padded box is for.
     */
    parity: (p, is) => {
      is.within(p.previewCellX("A panel: tinted"), p.wordCellX("A panel: tinted"), "1mm", "the panel's words start where Word starts them");
      is.within(p.previewCellX("Tinted, with no rule"), p.wordCellX("Tinted, with no rule"), "1mm", "and the tinted cell's, rule or no rule");
      is.within(p.previewCellY("A panel: tinted"), p.wordCellY("A panel: tinted"), "1mm", "on the line Word puts them on");
    },
  },
});
