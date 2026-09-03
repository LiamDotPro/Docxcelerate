import { Cell, Document, Paragraph, Row, Table, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { withBlocks } from "../_support/style.ts";

/**
 * How a cell's content sits against the height of a row taller than it.
 *
 * Every cell in a row is as tall as the tallest one, so the moment one cell
 * holds more than the others there is a decision to make about the rest. A
 * money figure beside a three-line description belongs on the *first* line of
 * it, not floating in the middle of a box — which is what `w:vAlign` decides
 * and what nobody notices until a row grows.
 *
 * The tall cell is made tall with three paragraphs rather than with one long
 * sentence, deliberately. A cell made tall by wrapping is as tall as its
 * column is wide, and Word autofits the columns (`tables/column-widths`) — so
 * the row would be one height on screen and another in Word, and the case
 * would be measuring that instead of the alignment it is about. Three
 * paragraphs are three lines in both engines whatever the column does.
 */
export default defineCase({
  id: "tables/cell-valign",
  feature: "table.cellValign",
  title: "Content sitting top, middle and bottom of a tall row",
  word: "Table Properties → Cell → Vertical alignment (w:vAlign)",
  claim: "supported",

  style: withBlocks({
    /** Held to the top of whatever height the row turns out to be. */
    top: { valign: "top" },
    /** Centred in it. */
    middle: { valign: "center" },
    /** Sat on its floor. */
    foot: { valign: "bottom" },
  }),

  document: template(
    <Document id="cell-valign" title="Cell vertical alignment">
      <Table
        id="rows"
        columns={[{ width: 70 }, { width: 30 }, { width: 30 }, { width: "auto" }]}
      >
        <Row id="head" header>
          <Cell id="h1">Description</Cell>
          <Cell id="h2">Top</Cell>
          <Cell id="h3">Middle</Cell>
          <Cell id="h4">Bottom</Cell>
        </Row>
        <Row id="tall">
          <Cell id="d1">
            <Paragraph id="d1a">A description</Paragraph>
            <Paragraph id="d1b">running to</Paragraph>
            <Paragraph id="d1c">three whole lines</Paragraph>
          </Cell>
          <Cell id="d2" variant="top">Up here</Cell>
          <Cell id="d3" variant="middle">Halfway</Cell>
          <Cell id="d4" variant="foot">Down there</Cell>
        </Row>
      </Table>
    </Document>
  ),

  regions: [
    { id: "top", anchor: "Up here" },
    { id: "middle", anchor: "Halfway" },
    { id: "bottom", anchor: "Down there" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.equal(a.cell("Up here").vAlign, "top", "a cell held to the top says so");
      is.equal(a.cell("Halfway").vAlign, "center", "one centred in the row says centre");
      is.equal(a.cell("Down there").vAlign, "bottom", "and one on the floor says bottom");
      is.equal(a.cell("A description").vAlign, null, "the cell that made the row tall says nothing");
      is.equal(a.cell("A description").paragraphs.length, 3, "and holds three paragraphs, which is what made it tall");
    },

    preview: (b, is) => {
      is.equal(b.cell("Up here").verticalAlign, "top", "the preview holds the first to the top");
      is.equal(b.cell("Halfway").verticalAlign, "middle", "centres the second");
      is.equal(b.cell("Down there").verticalAlign, "bottom", "and sits the third on the floor");

      // Where the words landed, which is what a computed style does not prove.
      // Three cells of one line each in a row three lines tall have to be in
      // three different places, and in this order.
      is.greater(b.cellTextTop("Halfway"), b.cellTextTop("Up here"), "the centred words are below the top ones");
      is.greater(b.cellTextTop("Down there"), b.cellTextTop("Halfway"), "and the bottom ones below those");
    },

    word: (c, is) => {
      is.equal(c.cell("Up here").vAlign, "top", "Word reads the first as top");
      is.equal(c.cell("Halfway").vAlign, "center", "the second as centred");
      is.equal(c.cell("Down there").vAlign, "bottom", "and the third as bottom");
      is.greater(c.cell("Halfway").y, c.cell("Up here").y, "and draws them in that order down the row");
      is.greater(c.cell("Down there").y, c.cell("Halfway").y, "top, middle, bottom");
    },

    /**
     * How far down the row each one landed, on screen against in Word. The
     * three tiers above all agree about which cell is centred; only this one
     * says the preview put it in the same place.
     */
    parity: (p, is) => {
      is.within(p.previewCellY("Up here"), p.wordCellY("Up here"), "1mm", "the top-aligned words sit where Word puts them");
      is.within(p.previewCellY("Halfway"), p.wordCellY("Halfway"), "1mm", "and the centred ones");
      is.within(p.previewCellY("Down there"), p.wordCellY("Down there"), "1mm", "and the ones on the floor");
    },
  },
});
