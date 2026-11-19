import { Document, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle } from "../_support/style.ts";

/**
 * A numbered list with a step under a step, and then a second list that
 * starts counting again.
 *
 * `list="number"` is what says it. A document that types its numbers has
 * numbers nobody counts: drop a step for one recipient and the rest keep the
 * figures they were written with. Numbering is the thing a renderer can do and
 * an author cannot, for the same reason a page number is. This case was
 * written red, with the prop already on its paragraphs and nothing reading it.
 *
 * The second list is the half that is easy to get wrong. Two lists that share
 * one numbering are one list — the second carries on at 3 — so each list has
 * to be a numbering of its own, starting at 1.
 *
 * **Each list is its own abstract numbering too, and that is for the
 * preview.** Word is content with two `w:num` over one `w:abstractNum`, each
 * with a start override; that is the form Word itself writes. docx-preview
 * keys its counters by the abstract one and keeps only the last `w:num` that
 * names it, so the earlier list is drawn with no numbers at all. Packing an
 * abstract numbering per list is the form both engines read — the same
 * decision a shape packed as VML rests on — and the assertion that the second
 * list has a marker at all is the one that holds it.
 *
 * The preview cannot be asked what number it reached: a CSS counter's value is
 * not in the DOM. What it can say is which counter each item counts with and
 * in what style, and two lists that count separately count with two. Figures
 * are the style a counter has when none is named, and Chrome reports the rule
 * without naming it — so "in figures" is read as no other style being there.
 */
export default defineCase({
  id: "lists/ordered-numbering",
  feature: "list.number",
  title: "Numbered steps, one nested, and a second list that starts again at one",
  word: "Home → Numbering (w:numPr, w:numFmt decimal / lowerLetter)",
  claim: "supported",

  style: caseStyle,

  document: template(
    <Document id="ordered-numbering" title="Numbered list">
      <Paragraph id="complete" list="number">Complete the form.</Paragraph>
      <Paragraph id="sign" list="number" listLevel={1}>Sign it on the last page.</Paragraph>
      <Paragraph id="return" list="number">Return it to us.</Paragraph>
      <Paragraph id="between">If you are applying jointly:</Paragraph>
      <Paragraph id="each" list="number">Each applicant signs.</Paragraph>
    </Document>
  ),

  regions: [
    { id: "first", anchor: "Complete the form" },
    { id: "nested", anchor: "Sign it on the last page" },
    { id: "second", anchor: "Return it to us" },
    { id: "between", anchor: "If you are applying jointly" },
    { id: "again", anchor: "Each applicant signs" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.equal(a.para("Complete the form").text, "Complete the form.", "a step's text is its words, with no number typed into it");

      is.equal(a.list("Complete the form").format, "decimal", "a step is numbered in figures");
      is.equal(a.list("Complete the form").text, "%1.", "followed by a full stop");
      is.equal(a.list("Sign it on the last page").format, "lowerLetter", "a step under a step is lettered");
      is.equal(a.list("Return it to us").numId, a.list("Complete the form").numId, "the next step counts in the same list");

      // The second list. A different numbering, starting at one, and — for the
      // preview's sake — a different abstract one behind it.
      is.equal(a.list("Each applicant signs").format, "decimal", "the second list is numbered too");
      is.notEqual(a.list("Each applicant signs").numId, a.list("Complete the form").numId, "in a numbering of its own");
      is.equal(a.list("Each applicant signs").start, 1, "that starts at one");
      is.notEqual(
        a.list("Each applicant signs").abstractNumId,
        a.list("Complete the form").abstractNumId,
        "over an abstract numbering of its own, which is what the preview counts by",
      );
    },

    preview: (b, is) => {
      const first = b.para("Complete the form");
      const nested = b.para("Sign it on the last page");
      const again = b.para("Each applicant signs");

      is.includes(first.marker, "counter(", "the preview counts a step");
      is.excludes(first.marker, "lower-", "in figures, the style a counter has when it names none");
      is.includes(nested.marker, "lower-alpha", "and letters the step under it");
      is.greater(nested.lines[0]?.x, first.lines[0]?.x ?? Infinity, "which is set further in");

      // The assertion that holds the abstract-numbering decision: a second
      // list sharing the first one's abstract numbering leaves the first with
      // no marker at all.
      is.includes(again.marker, "counter(", "the second list is counted as well");
      is.notEqual(again.marker, first.marker, "by a counter of its own, so it starts again at one");
    },

    word: (c, is) => {
      // The figures themselves are `ListFormat.ListString`, which the probe
      // cannot read yet. The hang is what it can.
      is.less(c.para("Complete the form").firstLineIndent, 0, "Word hangs a step's first line, where its number stands");
      is.greater(
        c.para("Sign it on the last page").leftIndent,
        c.para("Complete the form").leftIndent ?? Infinity,
        "and sets the step under it further in",
      );
    },
  },
});
