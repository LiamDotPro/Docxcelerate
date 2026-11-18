import { Document, Paragraph, template, Text } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * Two words in a sentence set bold, and the rest of it left alone.
 *
 * The term a notice turns on — the thirty days, the amount, the date — is the
 * thing a reader is meant to find first. A block style can make a whole
 * paragraph bold; `<Text bold>` is what makes part of one, and it reaches the
 * file as a run of its own carrying `w:b`, between two runs that do not.
 *
 * This case was written red, before there was a `<Text>` to write: the
 * sentence plain from end to end, and the assertions naming the two words
 * that should not have been.
 *
 * Each tier asks two things rather than one. A sentence packed as a single run
 * answers for every anchor inside it, so "the run holding these words is bold"
 * would be satisfied by a paragraph that is bold throughout — which a block
 * style could always do, and which is not this. The run has to *be* the two
 * words, and the words either side of it have to stay as they were.
 *
 * Word has no reading of a run, only of a range, and it answers for a range
 * that is bold in part with neither yes nor no: `wdUndefined`, which the probe
 * hands back as null. That is the fact asserted there — and why the same line
 * first checks the paragraph was found, since a paragraph that is missing
 * reads as null too.
 */
export default defineCase({
  id: "runs/bold-span",
  feature: "run.bold",
  title: "Two words in a sentence set bold",
  word: "Font → Bold, on a selection inside a paragraph (w:b on one w:r)",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="bold-span" title="Bold span">
      <Paragraph id="terms">
        Payment is due within <Text bold>thirty days</Text> of the date above.
      </Paragraph>
    </Document>
  ),

  regions: [
    { id: "sentence", anchor: "Payment is due" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.equal(
        a.para("Payment is due").text,
        "Payment is due within thirty days of the date above.",
        "the sentence is one paragraph, with every word of it in order",
      );
      is.notEqual(a.run("Payment is due").bold, true, "the words before the term are not bold");
      is.notEqual(a.run("of the date above").bold, true, "and nor are the words after it");

      // The two facts: the term is a run, and the run is what is bold.
      is.equal(a.run("thirty days").text, "thirty days", "the term is a run of its own");
      is.equal(a.run("thirty days").bold, true, "and that run is bold");
    },

    preview: (b, is) => {
      is.equal(b.para("Payment is due").lineCount, 1, "the sentence is still set on its one line");
      is.equal(b.run("Payment is due").fontWeight, "400", "the words before the term are drawn in the ordinary weight");

      is.equal(b.run("thirty days").text, "thirty days", "the preview draws the term as a run of its own");
      is.equal(b.run("thirty days").fontWeight, "700", "in bold");
    },

    word: (c, is) => {
      is.notEqual(c.para("Payment is due").missing, true, "Word finds the sentence");
      is.equal(c.para("Payment is due").lineCount, 1, "on its one line");

      // False for a sentence in one weight, true for one bold throughout.
      is.equal(c.para("Payment is due").bold, null, "and reads it as bold in part — neither bold nor not");
    },
  },
});
