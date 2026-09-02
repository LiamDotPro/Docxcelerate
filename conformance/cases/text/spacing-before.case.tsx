import { Document, Paragraph, template } from "docxcelerate/template";
import { defineCase } from "../../lib/case.mjs";
import { TWIPS_PER_PT, withBlocks } from "../_support/style.ts";

/**
 * The gap a paragraph leaves *above* itself.
 *
 * A heading has one — `DocumentTextBlockStyle` carries `spacingBeforePt`, and
 * the title and section heading both use it. A block does not: `DocumentBlockStyle`
 * has `spacingAfterPt` and nothing to match it, so a named block can close up
 * under itself but not open up above.
 *
 * That asymmetry costs more than it looks. Space above is how a block is
 * separated from what precedes it *without* the paragraph above having to know
 * something follows — which is the whole reason Word has both properties, and
 * why a document built only on space-after gets a stray gap at the bottom of
 * every page.
 */
export default defineCase({
  id: "text/spacing-before",
  feature: "paragraph.spacingBefore",
  title: "The space a block leaves above itself",
  word: "Paragraph → Spacing → Before (w:spacing w:before)",
  claim: "supported",

  style: withBlocks({
    /** A block set apart from whatever ran before it. */
    interlude: { spacingBeforePt: 18, spacingAfterPt: 6 },
  }),

  document: template(
    <Document id="spacing-before" title="Spacing before">
      <Paragraph id="a">Ordinary prose, leaving the document's gap under itself.</Paragraph>
      <Paragraph id="b" variant="interlude">Set apart from what came before it.</Paragraph>
      <Paragraph id="c">And back to ordinary prose underneath.</Paragraph>
    </Document>
  ),

  regions: [
    { id: "before", anchor: "Ordinary prose" },
    { id: "interlude", anchor: "Set apart" },
    { id: "after", anchor: "back to ordinary" },
  ],

  expect: {
    ooxml: (a, is) => {
      is.equal(a.para("Set apart").spacing?.before, 18 * TWIPS_PER_PT, "the block's 18pt above lands as 360 twips");
      is.equal(a.para("Set apart").spacing?.after, 6 * TWIPS_PER_PT, "and its 6pt under it still does too");
      is.equal(a.para("Ordinary prose").spacing?.before, null, "a paragraph that says nothing writes nothing");
    },

    preview: (b, is) => {
      // The larger of the two, not their sum — and that is Word's answer, not
      // a concession to CSS. Measured: the paragraph above leaves 10pt, this
      // block asks for 18pt above itself, and Word draws its top 33.3pt below
      // the previous paragraph's top, which is one 15.4pt line plus 18. So
      // Word collapses the pair exactly as adjacent CSS margins do, and the
      // preview happens to be right for a reason worth writing down rather
      // than rediscovering.
      is.within(b.gapAfter("Ordinary prose"), b.pt(18), "1pt", "the drawn gap is the larger of 10pt after and 18pt before");
      is.within(b.gapAfter("Set apart"), b.pt(6), "1pt", "and under the block, the larger of its 6pt and the next one's 0");
    },

    word: (c, is) => {
      is.equal(c.para("Set apart").spaceBefore, 18, "Word reads 18pt above");
      is.equal(c.para("Set apart").spaceAfter, 6, "and 6pt below");
    },

    parity: (p, is) => {
      is.within(p.previewY("Set apart"), p.wordY("Set apart"), "1mm", "the block starts where Word starts it");
      is.within(p.previewY("back to ordinary"), p.wordY("back to ordinary"), "1mm", "and what follows lands there too");
    },
  },
});
