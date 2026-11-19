import { Document, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { withBlocks } from "../_support/style.ts";

/**
 * A block's border drawn double, dashed or dotted.
 *
 * The preview widens a double border to three of the file's widths, because
 * CSS needs that to show two lines. That is a reading of Word's geometry from
 * documentation; the parity tier is what confirms it, and it needs Word.
 */
export default defineCase({
  id: "text/border-style",
  feature: "paragraph.borderStyle",
  title: "A paragraph bordered double, dashed or dotted",
  word: "Paragraph → Borders (w:pBdr w:val)",
  claim: "supported",

  style: withBlocks({
    doubleRule: { border: "111827", borderSides: ["bottom"], borderStyle: "double", paddingPt: 2 },
    tearOff: { border: "6B7280", borderSides: ["top"], borderStyle: "dashed", paddingPt: 6 },
    dotted: { border: "111827", borderStyle: "dotted", paddingPt: 4 },
  }),

  document: template(
    <Document id="border-style" title="Border style">
      <Paragraph id="a">Plain prose, with no border.</Paragraph>
      <Paragraph id="b" variant="doubleRule">Grand total, ruled twice beneath.</Paragraph>
      <Paragraph id="c">The paragraph after the double rule.</Paragraph>
      <Paragraph id="d" variant="tearOff">Detach and return with your payment.</Paragraph>
      <Paragraph id="e" variant="dotted">Dotted all round.</Paragraph>
    </Document>
  ),

  regions: [
    { id: "plain", anchor: "Plain prose" },
    { id: "double", anchor: "Grand total" },
    { id: "after", anchor: "The paragraph after" },
    { id: "tear", anchor: "Detach and return" },
    { id: "dotted", anchor: "Dotted all round" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.equal(a.para("Plain prose").pBdr, null, "prose with no border writes none");
      is.equal(a.para("Grand total").pBdr?.bottom?.style, "double", "the total is ruled twice");
      is.equal(a.para("Detach and return").pBdr?.top?.style, "dashed", "the tear-off line is dashed");
      is.equal(a.para("Dotted all round").pBdr?.left?.style, "dotted", "and the last is dotted");
    },

    preview: (b, is) => {
      is.equal(b.para("Grand total").borderBottomStyle, "double", "the preview draws the double rule");
      is.greater(b.para("Grand total").borderBottomWidth, 3, "wide enough to show two lines");
      is.equal(b.para("Detach and return").borderTopStyle, "dashed", "and the dashed line");
      is.equal(b.para("Dotted all round").borderTopStyle, "dotted", "and the dotted one");
    },

    word: (c, is) => {
      is.equal(c.para("Grand total").borders.bottom, true, "Word draws the double rule");
      is.equal(c.para("Detach and return").borders.top, true, "and the tear-off line");
    },

    parity: (p, is) => {
      is.within(p.previewY("The paragraph after"), p.wordY("The paragraph after"), "1mm", "the rows below a double rule sit where Word puts them");
    },
  },
});
