import { Document, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * A clause with its sub-clauses numbered (a), (b) and (i), and then a list
 * counted in capital roman numerals.
 *
 * A contract's sub-clause is `(a)`, not `a.`, and the default numerals cannot
 * say so. `listStyle` on the first item at a level picks that level's numerals
 * for the whole list, so the second sub-clause does not repeat it, and a
 * second list is its own numbering with its own choice.
 */
export default defineCase({
  id: "lists/number-style",
  feature: "list.numberStyle",
  title: "Sub-clauses numbered (a), (b), (i), and a list in roman numerals",
  word: "Home → Numbering → Define New Number Format (w:numFmt, w:lvlText)",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="number-style" title="Numbered clauses">
      <Paragraph id="shall" list="number">The Supplier shall:</Paragraph>
      <Paragraph id="deliver" list="number" listLevel={1} listStyle="lowerLetterParen">deliver the Goods;</Paragraph>
      <Paragraph id="date" list="number" listLevel={2} listStyle="lowerRomanParen">on the Delivery Date;</Paragraph>
      <Paragraph id="invoice" list="number" listLevel={1}>invoice monthly.</Paragraph>
      <Paragraph id="between">The Schedules:</Paragraph>
      <Paragraph id="schedule" list="number" listStyle="upperRoman">Pricing.</Paragraph>
    </Document>
  ),

  regions: [
    { id: "clause", anchor: "The Supplier shall" },
    { id: "deliver", anchor: "deliver the Goods" },
    { id: "date", anchor: "on the Delivery Date" },
    { id: "invoice", anchor: "invoice monthly" },
    { id: "schedule", anchor: "Pricing" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.equal(a.list("deliver the Goods").format, "lowerLetter", "a sub-clause is lettered");
      is.equal(a.list("deliver the Goods").text, "(%2)", "in brackets, with no stop");
      is.equal(a.list("invoice monthly").format, a.list("deliver the Goods").format, "the next sub-clause follows without saying so");
      is.equal(a.list("invoice monthly").numId, a.list("deliver the Goods").numId, "in the same list");
      is.equal(a.list("on the Delivery Date").format, "lowerRoman", "a clause under it is in roman numerals");
      is.equal(a.list("on the Delivery Date").text, "(%3)", "in brackets too");
      is.equal(a.list("Pricing").format, "upperRoman", "a second list chooses for itself");
      is.equal(a.list("Pricing").text, "%1.", "followed by a full stop");
    },

    preview: (b, is) => {
      is.includes(b.para("deliver the Goods").marker, "lower-alpha", "the preview counts a sub-clause in letters");
      is.includes(b.para("deliver the Goods").marker, "(", "inside brackets");
      is.includes(b.para("Pricing").marker, "upper-roman", "and the second list in capital roman numerals");
    },

    word: (c, is) => {
      is.less(c.para("deliver the Goods").firstLineIndent, 0, "Word hangs a sub-clause's first line, where its numeral stands");
      is.greater(
        c.para("on the Delivery Date").leftIndent,
        c.para("deliver the Goods").leftIndent ?? Infinity,
        "and sets the clause under it further in",
      );
    },
  },
});
