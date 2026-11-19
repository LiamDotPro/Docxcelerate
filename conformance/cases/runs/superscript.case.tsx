import { Document, Paragraph, template, Text } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * An ordinal and a chemical formula, set with raised and lowered letters.
 *
 * A look-alike such as `²` is a different character, not Word's Font →
 * Superscript. `<Text script>` reaches the file as a run of its own carrying
 * `w:vertAlign`, between two runs that do not.
 *
 * Each tier asks that the run is the raised words and that the words beside
 * it stay on the line. Word reads a range, not a run, so a paragraph raised in
 * part answers `wdUndefined`, which the probe hands back as null.
 */
export default defineCase({
  id: "runs/superscript",
  feature: "run.vertAlign",
  title: "Raised and lowered letters in a sentence",
  word: "Font → Superscript / Subscript, on a selection inside a paragraph (w:vertAlign on one w:r)",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="superscript" title="Raised and lowered">
      <Paragraph id="ordinal">
        Completed on the 1<Text script="superscript">st</Text> of March.
      </Paragraph>
      <Paragraph id="formula">
        Emissions of CO<Text script="subscript">2</Text> at the flue.
      </Paragraph>
    </Document>
  ),

  regions: [
    { id: "ordinal", anchor: "Completed on the" },
    { id: "formula", anchor: "Emissions of CO" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.equal(a.run("st").text, "st", "the raised letters are a run of their own");
      is.equal(a.run("st").vertAlign, "superscript", "and that run is raised");
      is.equal(a.run("2").vertAlign, "subscript", "the 2 in CO2 is lowered");
      is.equal(a.run("of March").vertAlign, null, "the words after it stay on the line");
    },

    preview: (b, is) => {
      is.equal(b.run("st").verticalAlign, "super", "the preview raises the letters");
      is.equal(b.run("2").verticalAlign, "sub", "and lowers the 2");
      is.equal(b.run("of March").verticalAlign, "baseline", "the words beside them stay on the line");
      is.equal(b.para("Completed on the").lineCount, 1, "the sentence is still on its one line");
    },

    word: (c, is) => {
      is.notEqual(c.para("Completed on the").missing, true, "Word finds the sentence");
      is.equal(c.para("Completed on the").superscript, null, "and reads it as raised in part");
      is.equal(c.para("Emissions of CO").subscript, null, "and the formula as lowered in part");
    },
  },
});
