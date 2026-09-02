import { Document, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { caseStyle, withBlocks } from "../_support/style.ts";

/**
 * Leading: the document's, and a block's own.
 *
 * Word writes leading as an exact number of twips with `w:lineRule="auto"`,
 * which means "this many twentieths of a point per line" — not a multiplier.
 * A browser writes it as a multiplier of the font size. The two agree only if
 * whoever converts between them uses the same font size the file did, and that
 * is exactly the kind of arithmetic that is right in one place and wrong in
 * another.
 *
 * A row of charges set at 1.2 against prose at 1.4 is a real invoice's real
 * difference, so both are here.
 */
export default defineCase({
  id: "text/line-height",
  feature: "paragraph.lineHeight",
  title: "Leading, from the document and from a block",
  word: "Paragraph → Spacing → Line spacing → Multiple (w:spacing w:line)",

  /**
   * Written right, drawn right, read right — and the two engines still place
   * heavily-led text about 1.6mm apart.
   *
   * The cause is where each engine puts the text inside a line taller than the
   * face needs. Word with an exact line height sits the baseline near the
   * bottom of it; CSS splits the leftover leading evenly above and below. At
   * ordinary settings that is nothing — measured, 1.1 leading drifts -0.39mm
   * and 1.4 drifts +0.16mm, both well inside the 1mm this suite holds a
   * paragraph to. At 2.0 it reaches 1.62mm, and it accumulates down the page.
   *
   * It is not closed by choosing a different rule; that was measured too, and
   * every alternative is worse (`auto` 2.29mm, `atLeast` 20.93mm, and 2.35mm
   * even with docx-preview's `atLeast` reading corrected by hand). Closing it
   * means making CSS distribute half-leading the way Word does, which CSS does
   * not expose. So it is recorded at its measured size rather than hidden
   * under a wider tolerance.
   */
  claim: "partial",
  knownRed: ["parity"],

  style: withBlocks({
    /** Set tighter than prose, the way a table row is. */
    tight: { lineHeight: 1.1 },
    /** And looser, for a standfirst that wants air. */
    airy: { lineHeight: 2 },
  }),

  document: template(
    <Document id="line-height" title="Line height">
      <Paragraph id="a">
        The document's own leading, over enough words to run to three lines on an A4 page
        with twenty millimetre margins, because leading is the distance between lines and a
        paragraph of one line has no lines to be distant from each other.
      </Paragraph>
      <Paragraph id="b" variant="tight">
        Set tighter by its block, over the same three lines, so the two can be compared by
        height rather than by the numbers each of them claims. This is what a row of charges
        is set at when prose around it is not.
      </Paragraph>
      <Paragraph id="c" variant="airy">
        Set looser by its block, again over three lines, which is the other direction and the
        one that catches a renderer treating any declared leading as a maximum rather than as
        the exact distance the file asked for.
      </Paragraph>
    </Document>
  ),

  regions: [
    { id: "default", anchor: "The document's own leading" },
    { id: "tight", anchor: "Set tighter by its block" },
    { id: "airy", anchor: "Set looser by its block" },
  ],

  expect: {
    ooxml: (a, is) => {
      const bodyPt = caseStyle.typography.bodySizePt;

      // `exact`, not `auto` and not `atLeast`. Measured over this very page,
      // as the distance between where the preview puts each paragraph and
      // where Word puts it: exact 1.62mm, auto 2.29mm, atLeast 20.93mm (and
      // 2.35mm even with docx-preview's atLeast reading corrected by hand).
      // `atLeast` is the kinder rule inside Word, because it grows a line
      // rather than clipping one — but Word grows it to the face's own natural
      // line height and CSS does not grow at all, so it is the rule the two
      // engines agree about least. Re-open this with a measurement, not an
      // opinion.
      is.equal(a.para("The document's own leading").spacing?.lineRule, "exact", "leading is written as an exact line height");
      is.within(
        a.para("The document's own leading").spacing?.line,
        bodyPt * caseStyle.typography.bodyLineHeight * 20,
        2,
        "1.4 of 11pt is 308 twips",
      );
      is.within(a.para("Set tighter by its block").spacing?.line, bodyPt * 1.1 * 20, 2, "the tight block writes 1.1 of 11pt");
      is.within(a.para("Set looser by its block").spacing?.line, bodyPt * 2 * 20, 2, "the airy block writes 2 of 11pt");
    },

    preview: (b, is) => {
      const bodyPx = b.pt(caseStyle.typography.bodySizePt);

      is.within(b.lineHeight("The document's own leading"), bodyPx * 1.4, "1px", "the preview leads the body at 1.4");
      is.within(b.lineHeight("Set tighter by its block"), bodyPx * 1.1, "1px", "and the tight block at 1.1");
      is.within(b.lineHeight("Set looser by its block"), bodyPx * 2, "1px", "and the airy block at 2");

      // The property is one thing; the drawn height is the thing a reader sees.
      is.equal(b.para("The document's own leading").lineCount, 3, "the default paragraph runs to three lines");
      is.greater(
        b.para("Set looser by its block").height,
        b.para("Set tighter by its block").height,
        "and a looser block is taller than a tighter one",
      );
    },

    word: (c, is) => {
      is.within(c.para("The document's own leading").lineSpacing, caseStyle.typography.bodySizePt * 1.4, "0.5pt", "Word reads the body leading");
      is.within(c.para("Set tighter by its block").lineSpacing, caseStyle.typography.bodySizePt * 1.1, "0.5pt", "and the tight block's");
    },

    parity: (p, is) => {
      is.within(p.previewY("Set tighter by its block"), p.wordY("Set tighter by its block"), "1mm", "the tight block starts where Word starts it");
      is.within(p.previewY("Set looser by its block"), p.wordY("Set looser by its block"), "1mm", "and so does the airy one, three lines later");
    },
  },
});
