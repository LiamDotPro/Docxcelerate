import { Document, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle, TWIPS_PER_PT, withBlocks } from "../_support/style.ts";

/**
 * The gap the document leaves under a paragraph, and under one that says its
 * own.
 *
 * This is the most-used paragraph property there is — every document sets it
 * once and then never thinks about it — which is exactly why it is worth
 * measuring rather than assuming. A preview that draws it in CSS margins and a
 * Word file that writes it in twips have two chances to disagree, and the gap
 * between paragraphs is the thing a reader notices first.
 */
export default defineCase({
  id: "text/spacing-after",
  feature: "paragraph.spacingAfter",
  title: "The space left under a paragraph, from the document and from a block",
  word: "Paragraph → Spacing → After (w:spacing w:after)",
  claim: "supported",

  style: withBlocks({
    /** A block that closes up under itself — a heading over its own list. */
    tight: { spacingAfterPt: 0 },
    /** And one that opens up, to prove the block wins over the document. */
    loose: { spacingAfterPt: 24 },
  }),

  document: template(
    <Document id="spacing-after" title="Spacing after">
      <Paragraph id="a">Default spacing under this one.</Paragraph>
      <Paragraph id="b" variant="tight">Closed up under this one.</Paragraph>
      <Paragraph id="c" variant="loose">Opened up under this one.</Paragraph>
      <Paragraph id="d">The last one, back to the default.</Paragraph>
    </Document>
  ),

  regions: [
    { id: "default", anchor: "Default spacing" },
    { id: "tight", anchor: "Closed up" },
    { id: "loose", anchor: "Opened up" },
    { id: "last", anchor: "The last one" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.equal(
        a.para("Default spacing").spacing?.after,
        caseStyle.paragraph.spacingAfterPt * TWIPS_PER_PT,
        "the document's 10pt lands as 200 twips",
      );
      is.equal(a.para("Closed up").spacing?.after, 0, "a block that says zero writes zero");
      is.equal(a.para("Opened up").spacing?.after, 24 * TWIPS_PER_PT, "a block's 24pt wins over the document's 10pt");
    },

    preview: (b, is) => {
      // The gap between two paragraphs is the distance from one's bottom to the
      // next one's top. Reading the CSS margin instead would measure what
      // docx-preview decided rather than what the reader is shown.
      is.within(b.gapAfter("Default spacing"), b.pt(10), "1pt", "the default gap is 10pt on screen");
      is.within(b.gapAfter("Closed up"), 0, "1pt", "the tight block leaves no gap");
      is.within(b.gapAfter("Opened up"), b.pt(24), "1pt", "the loose block leaves 24pt");
    },

    word: (c, is) => {
      is.equal(c.para("Default spacing").spaceAfter, 10, "Word reads the document's 10pt");
      is.equal(c.para("Closed up").spaceAfter, 0, "Word reads the tight block's zero");
      is.equal(c.para("Opened up").spaceAfter, 24, "Word reads the loose block's 24pt");
    },

    /** The screen and the file agree about where the next paragraph starts. */
    parity: (p, is) => {
      is.within(p.previewY("Closed up"), p.wordY("Closed up"), "1mm", "the tight paragraph starts where Word puts it");
      is.within(p.previewY("Opened up"), p.wordY("Opened up"), "1mm", "the loose paragraph starts where Word puts it");
      is.within(p.previewY("The last one"), p.wordY("The last one"), "1mm", "and so does everything under them");
    },
  },
});
