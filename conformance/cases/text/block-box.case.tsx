import { Document, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { withBlocks } from "../_support/style.ts";

/**
 * The box a block draws around a paragraph: fill, border, and the space between
 * the edge and the words.
 *
 * Word has no padding on a paragraph. What it has is `w:pBdr`'s `w:space`,
 * which is the gap between the border and the text — and only on the sides a
 * border is actually drawn on. So a tinted panel with no border has nothing to
 * hold its words off its own edge, and a renderer that maps padding onto
 * `w:space` will draw one thing on screen and another in Word.
 *
 * That asymmetry is the reason this case exists, and the reason its parity tier
 * matters more than its OOXML one.
 */
export default defineCase({
  id: "text/block-box",
  feature: "paragraph.blockBox",
  title: "A paragraph drawn as a filled, bordered box",
  word: "Paragraph → Borders and Shading (w:shd, w:pBdr)",
  claim: "supported",

  style: withBlocks({
    /** A tinted panel, bordered all round. */
    panel: {
      fill: "EEF2FF",
      border: "C7D2FE",
      borderWidthPt: 1,
      paddingPt: 8,
    },
    /** A note with a rule under it and nothing else. */
    underlined: {
      border: "D1D5DB",
      borderSides: ["bottom"],
      paddingPt: 4,
    },
    /** Fill with no border at all — the case Word cannot pad. */
    tinted: {
      fill: "FEF3C7",
      paddingPt: 6,
    },
  }),

  document: template(
    <Document id="block-box" title="Block box">
      <Paragraph id="a">Plain prose, in no box at all.</Paragraph>
      <Paragraph id="b" variant="panel">A panel: tinted, bordered on all four sides, and padded.</Paragraph>
      <Paragraph id="c" variant="underlined">A note with a rule beneath it.</Paragraph>
      <Paragraph id="d" variant="tinted">Tinted, with no border to hold the words off its edge.</Paragraph>
    </Document>
  ),

  regions: [
    { id: "plain", anchor: "Plain prose" },
    { id: "panel", anchor: "A panel: tinted" },
    { id: "underlined", anchor: "A note with a rule" },
    { id: "tinted", anchor: "Tinted, with no border" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.equal(a.para("Plain prose").shd, null, "prose in no box writes no shading");
      is.equal(a.para("Plain prose").pBdr, null, "and no borders");

      is.equal(a.para("A panel: tinted").shd, "EEF2FF", "the panel's fill is written as shading");
      is.equal(a.para("A panel: tinted").pBdr?.top?.color, "C7D2FE", "and its border is drawn on top");
      is.equal(a.para("A panel: tinted").pBdr?.bottom?.color, "C7D2FE", "and bottom");
      is.equal(a.para("A panel: tinted").pBdr?.left?.color, "C7D2FE", "and left");
      is.equal(a.para("A panel: tinted").pBdr?.right?.color, "C7D2FE", "and right");
      is.equal(a.para("A panel: tinted").pBdr?.top?.spacePt, 8, "with 8pt between the border and the words");

      is.equal(a.para("A note with a rule").pBdr?.bottom?.color, "D1D5DB", "the note is ruled underneath");
      is.equal(a.para("A note with a rule").pBdr?.top, null, "and nowhere else");
      is.equal(a.para("A note with a rule").pBdr?.left, null, "nowhere else at all");

      is.equal(a.para("Tinted, with no border").shd, "FEF3C7", "the tinted block is filled");
    },

    preview: (b, is) => {
      is.equal(b.para("A panel: tinted").background, b.hex("EEF2FF"), "the panel draws its fill");
      is.equal(b.para("A panel: tinted").borderTopColor, b.hex("C7D2FE"), "and its border");
      is.greater(b.para("A panel: tinted").paddingLeft, 0, "with the words held off its left edge");
      is.equal(b.para("Tinted, with no border").background, b.hex("FEF3C7"), "the tinted block draws its fill");
    },

    word: (c, is) => {
      is.equal(c.para("A panel: tinted").shading, "EEF2FF", "Word fills the panel");
      is.equal(c.para("A panel: tinted").borders.top, true, "and draws its border on every side");
      is.equal(c.para("A panel: tinted").borders.bottom, true, "every side");
      is.equal(c.para("A note with a rule").borders.bottom, true, "the note is ruled beneath");
      is.equal(c.para("A note with a rule").borders.top, false, "and not above");
    },

    /**
     * Where the words sit inside the box, on screen against in Word.
     *
     * This is the assertion the case is really for. A padded panel whose text
     * starts 8pt in on screen and 0pt in inside Word is a preview that lied
     * about the one thing a box is for.
     */
    parity: (p, is) => {
      is.within(p.previewX("A panel: tinted"), p.wordX("A panel: tinted"), "1mm", "the panel's words start where Word starts them");
      is.within(p.previewY("A panel: tinted"), p.wordY("A panel: tinted"), "1mm", "and on the line Word puts them on");
      is.within(p.previewX("Tinted, with no border"), p.wordX("Tinted, with no border"), "1mm", "and so do the tinted block's, border or no border");
      is.within(p.previewY("Tinted, with no border"), p.wordY("Tinted, with no border"), "1mm", "on the line Word puts it on");
    },
  },
});
