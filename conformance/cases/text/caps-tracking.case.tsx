import { Document, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { withBlocks } from "../_support/style.ts";

/**
 * A label set in capitals, opened up so it can be read as one.
 *
 * Capitals set at a text size are set at the wrong spacing — the letterforms
 * were drawn to sit under lower case — so a small label in caps needs tracking
 * or it reads as a smudge. Word writes the casing as `w:caps`, which leaves the
 * text alone and prints it differently, and the tracking as `w:spacing` on the
 * run, in twips.
 *
 * Both are places a preview can quietly diverge: `w:caps` is a *rendering*
 * instruction, so anything reading the text back sees the original case, and a
 * tracking of 0.1em on a 7pt face is 14 twips — small enough that getting the
 * conversion wrong looks like nothing at all until the label wraps.
 */
export default defineCase({
  id: "text/caps-tracking",
  feature: "paragraph.capsTracking",
  title: "Small capitals, opened up with tracking",
  word: "Font → All caps + Advanced → Character Spacing (w:caps, w:spacing)",
  claim: "supported",

  style: withBlocks({
    /** The label over a field: small, capitalised, and tracked open. */
    label: {
      fontSizePt: 7,
      weight: "bold",
      transform: "uppercase",
      letterSpacingEm: 0.12,
      color: "6B7280",
      spacingAfterPt: 2,
    },
  }),

  document: template(
    <Document id="caps-tracking" title="Caps and tracking">
      <Paragraph id="a" variant="label">Invoice reference</Paragraph>
      <Paragraph id="b">INV-2026-0142</Paragraph>
    </Document>
  ),

  regions: [
    { id: "label", anchor: "Invoice reference" },
    { id: "value", anchor: "INV-2026-0142" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.equal(a.para("Invoice reference").run.caps, true, "the label is capitalised by w:caps");
      is.equal(
        a.para("Invoice reference").text,
        "Invoice reference",
        "and the text in the file is still what the document wrote — caps is how it prints, not what it says",
      );
      is.equal(a.para("Invoice reference").run.szHalfPt, 14, "the label is 7pt");
      // 0.12em of 7pt is 0.84pt, which Word counts as 17 twips.
      is.within(a.para("Invoice reference").run.spacing, 0.12 * 7 * 20, 1, "tracked open by 0.12em of its own size");
    },

    preview: (b, is) => {
      is.equal(b.para("Invoice reference").textTransform, "uppercase", "the preview capitalises it");
      is.within(b.para("Invoice reference").letterSpacing, b.pt(0.12 * 7), "0.3px", "and tracks it by the same amount");
      is.within(b.para("Invoice reference").fontSize, b.pt(7), "0.5px", "at the same size");
    },

    word: (c, is) => {
      is.equal(c.para("Invoice reference").allCaps, true, "Word prints the label in capitals");
      is.within(c.para("Invoice reference").characterSpacing, 0.12 * 7, "0.1pt", "tracked by 0.84pt");
    },

    /**
     * Tracking changes how wide a line is, which is the fact worth comparing:
     * two engines agreeing on the number and disagreeing on the drawn width
     * would mean one of them applied it to the wrong thing.
     */
    parity: (p, is) => {
      is.within(p.previewWidth("Invoice reference"), p.wordWidth("Invoice reference"), "2mm", "the tracked label is the same width in both");
    },
  },
});
