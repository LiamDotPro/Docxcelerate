import { Document, Paragraph, template, Text } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { withBlocks } from "../_support/style.ts";

/**
 * An amount in a sentence drawn from a theme block: tinted, coloured, bold.
 *
 * `<Text bold>` says a phrase is stressed; `<Text variant>` says which of the
 * theme's named settings it is drawn in, so the colours stay the theme's to
 * change. The block's `color` and `fill` reach the file as `w:color` and
 * `w:shd` on one run, between two runs that carry neither.
 *
 * Word has no reading of a run, only of a range, so the Word tier asks only
 * that the sentence is found and set on one line, as `runs/bold-span` does.
 */
export default defineCase({
  id: "runs/variant-span",
  feature: "run.variant",
  title: "An amount in a sentence drawn from a theme block",
  word: "Font colour + Shading on a selection (w:color, w:shd on one w:r)",
  claim: "supported",

  style: withBlocks({ amountDue: { fill: "FFF3C4", color: "92400E", weight: "bold" } }),

  document: template(
    <Document id="variant-span" title="Variant span">
      <Paragraph id="due">
        The amount due is <Text variant="amountDue">£1,240.00</Text> by 30 November.
      </Paragraph>
    </Document>
  ),

  regions: [
    { id: "sentence", anchor: "The amount due" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.equal(a.run("£1,240.00").text, "£1,240.00", "the amount is a run of its own");
      is.equal(a.run("£1,240.00").shd, "FFF3C4", "shaded behind the words");
      is.equal(a.run("£1,240.00").color, "92400E", "in the block's colour");
      is.equal(a.run("£1,240.00").bold, true, "and its weight");
      is.equal(a.run("The amount due").shd, null, "the words before it are not shaded");
    },

    preview: (b, is) => {
      is.equal(b.para("The amount due").lineCount, 1, "the sentence is still set on its one line");
      is.equal(b.run("£1,240.00").background, "rgb(255, 243, 196)", "the preview shades the amount");
      is.equal(b.run("£1,240.00").color, "rgb(146, 64, 14)", "in the block's colour");
      is.equal(b.run("£1,240.00").fontWeight, "700", "and weight");
    },

    word: (c, is) => {
      is.notEqual(c.para("The amount due").missing, true, "Word finds the sentence");
      is.equal(c.para("The amount due").lineCount, 1, "on its one line");
    },
  },
});
