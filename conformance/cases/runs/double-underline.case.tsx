import { Document, Paragraph, template, Text } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * The accounting double rule under a final figure, and a single rule beside it.
 *
 * A paragraph border would rule the whole line; `<Text underline="double">`
 * reaches the file as a run of its own carrying `w:u w:val="double"`, and the
 * words before it carry no rule at all.
 *
 * There is no Word reading of `Font.Underline` in the probe, and no Word tier on
 * this Mac, so Word is asked only that the paragraph is found and is one line.
 */
export default defineCase({
  id: "runs/double-underline",
  feature: "run.underline.double",
  title: "A double rule under a total",
  word: "Font → Underline style → Double, on a selection inside a paragraph (w:u w:val=\"double\" on one w:r)",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="double-underline" title="Double underline">
      <Paragraph id="total">
        Total due <Text bold underline="double">£1,240.00</Text>
      </Paragraph>
      <Paragraph id="terms">
        Read the <Text underline>terms</Text> first.
      </Paragraph>
    </Document>
  ),

  regions: [
    { id: "total", anchor: "Total due" },
    { id: "terms", anchor: "Read the" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.equal(a.run("£1,240.00").text, "£1,240.00", "the figure is a run of its own");
      is.equal(a.run("£1,240.00").underline, "double", "and that run carries the double rule");
      is.equal(a.run("Total due").underline, null, "the words before it carry no rule");
      is.equal(a.run("terms").underline, "single", "underline alone is still the single rule");
    },

    preview: (b, is) => {
      is.equal(b.run("£1,240.00").textDecoration, "underline", "the preview underlines the figure");
      is.equal(b.run("£1,240.00").textDecorationStyle, "double", "with a double rule");
      is.equal(b.run("terms").textDecorationStyle, "solid", "and the single rule stays solid");
      is.equal(b.run("Total due").textDecoration, "none", "the words before the figure have no rule");
      is.equal(b.para("Total due").lineCount, 1, "the sentence is still on its one line");
    },

    word: (c, is) => {
      is.notEqual(c.para("Total due").missing, true, "Word finds the sentence");
      is.equal(c.para("Total due").lineCount, 1, "and sets it on one line");
    },
  },
});
