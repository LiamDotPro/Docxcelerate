import { Document, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { COLUMN_MM, TWIPS_PER_MM, withBlocks } from "../_support/style.ts";

/**
 * A paragraph set in from the margin, and one whose first line is.
 *
 * Until this landed, the only indent the framework wrote was the negative one
 * a bleeding block uses to reach past the margin: `blockIndent` existed to
 * escape the text column, not to sit inside it. So a quotation could not be
 * inset, a continuation could not hang, and a first-line indent could not be
 * set at all.
 *
 * All three go on the block rather than the node: an indent is what a named
 * thing looks like — a `quote` is inset because it is a quote — and a theme
 * that swapped 8mm for 12mm should not have to touch a paragraph to do it.
 */
export default defineCase({
  id: "text/indent",
  feature: "paragraph.indent",
  title: "A paragraph inset from the margin, hung, or indented on its first line",
  word: "Paragraph → Indentation → Left / Right / Special (w:ind)",
  claim: "supported",

  style: withBlocks({
    /** Inset from both sides, the way a pulled quotation is. */
    quote: { indentMm: 10, indentRightMm: 10 },
    /** Every paragraph after the first in a book. */
    firstLine: { firstLineIndentMm: 8 },
    /** A definition whose continuation lines clear its term. */
    hanging: { indentMm: 12, hangingIndentMm: 12 },
  }),

  document: template(
    <Document id="indent" title="Indentation">
      <Paragraph id="a">Flush to both margins, which is what a paragraph does when it says nothing.</Paragraph>
      <Paragraph id="b" variant="quote">
        Inset from both margins, the way a pulled quotation sits — far enough in that the
        eye reads it as a different voice, and long enough here to wrap onto a second line
        so both edges can be measured.
      </Paragraph>
      <Paragraph id="c" variant="firstLine">
        Indented on the first line only, which is how a book marks a new paragraph without
        leaving a gap above it, and which needs two lines before the difference between the
        first and the rest is visible at all.
      </Paragraph>
      <Paragraph id="d" variant="hanging">
        Hanging, so that the first line starts at the margin and every line under it clears
        the width of the term it belongs to.
      </Paragraph>
    </Document>
  ),

  regions: [
    { id: "flush", anchor: "Flush to both margins" },
    { id: "quote", anchor: "Inset from both margins" },
    { id: "firstLine", anchor: "Indented on the first line" },
    { id: "hanging", anchor: "Hanging, so that" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.equal(a.para("Flush to both margins").ind, null, "an unindented paragraph writes no w:ind");

      is.within(a.para("Inset from both margins").ind?.left, 10 * TWIPS_PER_MM, 2, "the quote is inset 10mm on the left");
      is.within(a.para("Inset from both margins").ind?.right, 10 * TWIPS_PER_MM, 2, "and 10mm on the right");

      is.within(a.para("Indented on the first line").ind?.firstLine, 8 * TWIPS_PER_MM, 2, "the first line is indented 8mm");
      is.equal(a.para("Indented on the first line").ind?.left, null, "and the block itself is not");

      is.within(a.para("Hanging, so that").ind?.left, 12 * TWIPS_PER_MM, 2, "the hanging block sits 12mm in");
      is.within(a.para("Hanging, so that").ind?.hanging, 12 * TWIPS_PER_MM, 2, "with its first line pulled back the same distance");
    },

    preview: (b, is) => {
      is.within(b.textLeft("Inset from both margins"), b.mm(10), "1mm", "the quote draws 10mm in");
      is.within(b.textRight("Inset from both margins"), b.mm(COLUMN_MM - 10), "1mm", "and stops 10mm short of the right margin");

      // A first-line indent is only itself if the *first* line moved and the
      // rest did not, so both are measured rather than the box.
      is.within(b.firstLineLeft("Indented on the first line"), b.mm(8), "1mm", "the first line starts 8mm in");
      is.within(b.textLeft("Indented on the first line"), 0, "1mm", "and the lines under it start at the margin");

      is.within(b.firstLineLeft("Hanging, so that"), 0, "1mm", "the hanging block's first line starts at the margin");
      is.within(b.textLeft("Hanging, so that"), b.mm(12), "1mm", "and its later lines clear 12mm");
    },

    word: (c, is) => {
      is.within(c.para("Inset from both margins").leftIndent, c.mm(10), "0.5mm", "Word reads a 10mm left indent");
      is.within(c.para("Inset from both margins").rightIndent, c.mm(10), "0.5mm", "and a 10mm right one");
      is.within(c.para("Indented on the first line").firstLineIndent, c.mm(8), "0.5mm", "Word reads the first-line indent");
      is.within(c.para("Hanging, so that").firstLineIndent, -c.mm(12), "0.5mm", "and reads a hang as a negative first line");
    },

    parity: (p, is) => {
      is.within(p.previewX("Inset from both margins"), p.wordX("Inset from both margins"), "1mm", "the quote starts where Word starts it");
      is.within(p.previewX("Indented on the first line"), p.wordX("Indented on the first line"), "1mm", "and so does the first-line indent");
    },
  },
});
