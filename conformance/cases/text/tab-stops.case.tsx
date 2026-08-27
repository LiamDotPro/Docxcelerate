import { Document, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { COLUMN_MM, TWIPS_PER_MM, withBlocks } from "../_support/style.ts";

/**
 * Two things on one line, the second one aligned to a stop rather than to a
 * run of spaces.
 *
 * A contents line with a dot leader, a signature block with a date on the
 * right, a price beside a description — all of them are one paragraph with a
 * tab in it, and none of them can be written today. The workaround is a table,
 * which draws a grid where a line was wanted and cannot be justified or wrapped.
 *
 * The stops go on the block: a `contentsLine` has the same stop wherever it
 * appears, and a theme that moves the leader from 150mm to 160mm should not
 * have to visit every line to do it.
 */
export default defineCase({
  id: "text/tab-stops",
  feature: "paragraph.tabStops",
  title: "A tab stop, with and without a leader",
  word: "Paragraph → Tabs (w:tabs)",

  /**
   * Written correctly and drawn correctly by Word; invisible in the preview.
   *
   * docx-preview has nowhere to put a tab stop — HTML has no tab-stop model,
   * and a `\t` in a text node collapses to a single space. So a contents line
   * that Word sets with its page number on the right margin and a dot leader
   * running up to it draws on screen as one run of words with a space in it.
   *
   * That is recorded rather than papered over. Settling the preview means
   * writing back a fact the file declares and Word draws; laying out tab stops
   * in CSS would be inventing a layout engine, and a preview that invents is
   * the thing this whole suite exists to prevent.
   */
  claim: "partial",
  knownRed: ["preview", "parity"],

  style: withBlocks({
    /** A contents line: text, dots, then a page number on the right margin. */
    contentsLine: {
      tabStopsMm: [{ at: COLUMN_MM, align: "right", leader: "dot" }],
      spacingAfterPt: 2,
    },
    /** A signature block: a name on the left, a date halfway across. */
    signature: {
      tabStopsMm: [{ at: 90 }],
    },
  }),

  document: template(
    <Document id="tab-stops" title="Tab stops">
      <Paragraph id="a" variant="contentsLine">{"What the agreement covers\t3"}</Paragraph>
      <Paragraph id="b" variant="contentsLine">{"When either party may end it\t11"}</Paragraph>
      <Paragraph id="c" variant="signature">{"Signed for the supplier\tDate"}</Paragraph>
    </Document>
  ),

  regions: [
    { id: "contents", anchor: "What the agreement covers" },
    { id: "signature", anchor: "Signed for the supplier" },
  ],

  expect: {
    ooxml: (a, is) => {
      const stops = a.para("What the agreement covers").tabs;

      is.equal(stops.length, 1, "the contents line declares one stop");
      is.within(stops[0]?.pos, COLUMN_MM * TWIPS_PER_MM, 2, "at the right margin");
      is.equal(stops[0]?.val, "right", "as a right stop");
      is.equal(stops[0]?.leader, "dot", "with a dot leader");

      is.within(a.para("Signed for the supplier").tabs[0]?.pos, 90 * TWIPS_PER_MM, 2, "the signature stop is 90mm in");
      is.equal(a.para("Signed for the supplier").tabs[0]?.val, "left", "and is a plain left stop");

      // The tab has to survive into the run as a tab, not as whitespace: a
      // paragraph whose text was flattened would look right and align to
      // nothing.
      is.includes(a.para("What the agreement covers").xml, "<w:tab/>", "the tab is written as a tab, not as spaces");
    },

    preview: (b, is) => {
      // What a stop is for: the number ends where the stop is, whatever the
      // description before it happens to be.
      is.within(b.textRight("What the agreement covers"), b.mm(COLUMN_MM), "2mm", "the page number lands on the right margin");
      is.within(b.textRight("When either party may end it"), b.mm(COLUMN_MM), "2mm", "and so does the next line's, despite a longer description");
      is.equal(b.para("What the agreement covers").lineCount, 1, "a contents line is one line");
    },

    word: (c, is) => {
      // The document's own stop, not Word's. Word's TabStops collection also
      // carries the default grid it puts every half inch across the page, so
      // a paragraph declaring one stop reports four — counting the collection
      // would be counting the page rather than the file.
      const stops = c.customTabStops("What the agreement covers");

      is.equal(stops.length, 1, "Word reads one stop of the document's own");
      is.within(stops[0]?.position, c.mm(COLUMN_MM), "0.5mm", "at the right margin");
      is.equal(stops[0]?.alignment, "right", "as a right stop");
      is.equal(stops[0]?.leader, "dot", "with a dot leader");

      is.equal(c.customTabStops("Signed for the supplier").length, 1, "and one on the signature line");
      is.within(c.customTabStops("Signed for the supplier")[0]?.position, c.mm(90), "0.5mm", "90mm in");
    },

    parity: (p, is) => {
      is.within(
        p.previewRight("What the agreement covers"),
        p.wordRight("What the agreement covers"),
        "2mm",
        "the tabbed number ends where Word ends it",
      );
    },
  },
});
